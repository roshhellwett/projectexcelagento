import { runInvariants } from './invariants.js';
import { issue, previewForTransition, transitionResult } from './operation-utils.js';
import {
  computeReconciliation,
  reconciliationAllocatedCells,
  reconcileSheetsArgsSchema,
  RECONCILIATION_LIMITS,
  RECONCILIATION_OUTCOMES,
  type ComputedReconciliation,
  type ReconciliationNumericTotal,
  type ReconciliationRow,
  type ReconcileSheetsArgs,
} from './reconciliation.js';
import type {
  Cell,
  CellRange,
  CellValue,
  InvariantResult,
  Operation,
  Sheet,
  ValidationResult,
  Workbook,
  WorkbookArtifact,
} from './types.js';
import { cellValueEquals, createCell, indexToColumn } from './workbook.js';

type ReportRole = WorkbookArtifact['sheets'][number]['role'];
const ROLES: ReportRole[] = ['summary', 'matched', 'exceptions', 'methodology'];
const RECORD_HEADERS = [
  'Outcome',
  'Side',
  'Source sheet',
  'Source row',
  'Source range',
  'Key cells',
  'Typed source key',
  'Normalized framed key',
  'Amount cell',
  'Amount',
  'Amount status',
  'Numeric text conversion',
  'Counterpart sheet',
  'Counterpart row',
  'Counterpart range',
  'Left minus right',
  'Delta status',
  'Detail',
];

interface ReportPlan {
  computed: ComputedReconciliation;
  names: Record<ReportRole, string>;
  summary: CellValue[][];
  methodology: CellValue[][];
  ranges: CellRange[];
  generatedCells: number;
  artifact: WorkbookArtifact;
}

/** Reserve space for each role and each collision suffix before truncating to Excel's limit. */
function reportNames(workbook: Workbook, prefix: string): Record<ReportRole, string> {
  const safePrefix =
    [...prefix]
      .map((character) =>
        character.charCodeAt(0) < 32 || ':\\/?*[]'.includes(character) ? '_' : character,
      )
      .join('')
      .trim()
      .replace(/^'+|'+$/g, '') || 'Reconciliation';
  const used = new Set(workbook.sheets.map((sheet) => sheet.name.toLowerCase()));
  const names = {} as Record<ReportRole, string>;
  for (const role of ROLES) {
    const label = role[0]!.toUpperCase() + role.slice(1);
    const desired = `${safePrefix.slice(0, 30 - label.length)} ${label}`;
    let name = desired;
    let suffixNumber = 2;
    while (used.has(name.toLowerCase())) {
      const suffix = ` (${suffixNumber++})`;
      name = desired.slice(0, 31 - suffix.length) + suffix;
    }
    used.add(name.toLowerCase());
    names[role] = name;
  }
  return names;
}

function qualifiedRange(range: CellRange): string {
  return `'${range.sheet.replace(/'/g, "''")}'!${range.startColumn}${range.startRow}:${range.endColumn}${range.endRow}`;
}
function qualifiedCell(sheet: string, column: string, row: number): string {
  return `'${sheet.replace(/'/g, "''")}'!${column}${row}`;
}
function typedSourceKey(row: ReconciliationRow): string {
  return JSON.stringify(
    row.keyValues.map((value) => {
      if (value instanceof Date)
        return ['date', Number.isFinite(value.getTime()) ? value.toISOString() : 'invalid_date'];
      if (typeof value === 'number')
        return [
          'number',
          Number.isFinite(value) ? (Object.is(value, -0) ? '-0' : value) : 'non_finite',
        ];
      return [value === null ? 'blank' : typeof value, value];
    }),
  );
}
function recordValues(row: ReconciliationRow): CellValue[] {
  return [
    row.outcome,
    row.side,
    row.sheet,
    row.sourceRow,
    qualifiedRange(row.sourceRange),
    row.keyCells.map((cell) => qualifiedCell(cell.sheet, cell.column, cell.row)).join('; '),
    typedSourceKey(row),
    row.key,
    row.amountCell === null
      ? null
      : qualifiedCell(row.amountCell.sheet, row.amountCell.column, row.amountCell.row),
    row.amount?.value ?? null,
    row.amount === null ? 'not_configured' : (row.amount.reason ?? 'numeric'),
    row.amount?.numericTextConverted ?? false,
    row.counterpart?.sheet ?? null,
    row.counterpart?.sourceRow ?? null,
    row.counterpart === null ? null : qualifiedRange(row.counterpart.sourceRange),
    row.delta,
    row.deltaReason ?? 'finite',
    row.detail,
  ];
}
function totalText(total: ReconciliationNumericTotal): string {
  return total.value === null ? 'Unavailable (overflow)' : String(total.value);
}

function artifactFor(
  computed: ComputedReconciliation,
  names: Record<ReportRole, string>,
): WorkbookArtifact {
  const { counts, amounts } = computed;
  const invalidKeys = counts.left.outcomes.invalid_key + counts.right.outcomes.invalid_key;
  const excludedAmounts =
    amounts === null ? 0 : amounts.left.total.excludedCount + amounts.right.total.excludedCount;
  const conversions =
    amounts === null
      ? 0
      : amounts.left.total.numericTextConversions + amounts.right.total.numericTextConversions;
  const proofs = amounts === null ? [] : [amounts.left.proof, amounts.right.proof];
  const numericComplete = proofs.every(
    (proof) => proof.countsBalanced && proof.numericCountsBalanced && proof.exclusionsBalanced,
  );
  const totalsStatus = proofs.some((proof) => proof.totalsBalanced === false)
    ? 'failed'
    : proofs.some((proof) => proof.totalsBalanced === null)
      ? 'warning'
      : 'passed';
  const discrepancies =
    counts.pairs.amount_mismatch +
    counts.left.outcomes.left_only +
    counts.right.outcomes.right_only;
  const overflow = computed.warnings.some((warning) => warning.code === 'numeric-overflow');
  return {
    id: `reconciliation:${names.summary}`,
    kind: 'reconciliation',
    title: `${names.summary} report`,
    sheets: ROLES.map((role) => ({ name: names[role], role })),
    sources: computed.sources.map((range) => ({ ...range })),
    facts: [
      { label: 'Left source rows', value: String(counts.left.sourceRows) },
      { label: 'Right source rows', value: String(counts.right.sourceRows) },
      { label: 'Matched one-to-one pairs', value: String(counts.pairs.matched) },
      { label: 'Exception source rows', value: String(counts.exceptionRows) },
      { label: 'Duplicate key groups', value: String(counts.duplicateGroups) },
      { label: 'Excluded amount cells', value: String(excludedAmounts) },
      { label: 'Numeric text conversions', value: String(conversions) },
      {
        label: 'Left numeric control total',
        value: amounts === null ? 'Not configured' : totalText(amounts.left.total),
      },
      {
        label: 'Right numeric control total',
        value: amounts === null ? 'Not configured' : totalText(amounts.right.total),
      },
    ],
    checks: [
      {
        id: 'source_preservation',
        label: 'Original workbook preserved',
        status: 'passed',
        detail:
          'Exactly four value-only sheets appended. Original sheet order, source cells, formulas, formats, and date epoch are checked by an explicit preservation invariant.',
      },
      {
        id: 'row_accounting',
        label: 'Complete source-row accounting',
        status: counts.rowAccountingComplete ? 'passed' : 'failed',
        detail: `${counts.left.accountedRows}/${counts.left.sourceRows} left and ${counts.right.accountedRows}/${counts.right.sourceRows} right rows classified exactly once, including all-blank rows.`,
      },
      {
        id: 'numeric_accounting',
        label: 'Numeric inclusions and exclusions accounted',
        status: numericComplete ? 'passed' : 'failed',
        detail:
          amounts === null
            ? 'Amounts not configured.'
            : 'Each side: numeric + excluded = source rows; matched + exception populations partition both numeric and excluded counts.',
      },
      {
        id: 'control_totals',
        label: 'Control-total partition proof',
        status: totalsStatus,
        detail:
          amounts === null
            ? 'Amounts not configured.'
            : 'Total = matched + exceptions within the reported accumulation roundoff bound. Overflow makes this proof unavailable, not a fabricated pass.',
      },
      {
        id: 'unique_keys',
        label: 'Unambiguous pairing',
        status: counts.duplicateGroups ? 'warning' : 'passed',
        detail: `${counts.duplicateGroups} duplicate group(s); every row of a duplicate group is an exception and none is paired.`,
      },
      {
        id: 'key_quality',
        label: 'Key-value quality',
        status: invalidKeys ? 'warning' : 'passed',
        detail: `${invalidKeys} invalid key row(s). Blank, invalid, formula error, unsupported, and volatile values are explicitly excluded from pairing.`,
      },
      {
        id: 'amount_quality',
        label: 'Amount-value quality',
        status: excludedAmounts || conversions || overflow ? 'warning' : 'passed',
        detail: `${excludedAmounts} excluded amount(s), ${conversions} conservative numeric-text conversion(s). ${overflow ? 'Finite-range overflow is explicitly reported.' : 'Invalid amounts are never treated as zero.'}`,
      },
      {
        id: 'business_discrepancies',
        label: 'Data discrepancies (not an accounting failure)',
        status: discrepancies || counts.pairs.invalid_amount ? 'warning' : 'passed',
        detail: `${counts.pairs.amount_mismatch} amount-mismatch pair(s), ${counts.pairs.invalid_amount} invalid-amount pair(s), ${counts.left.outcomes.left_only} left-only and ${counts.right.outcomes.right_only} right-only rows. No financial conclusion is inferred.`,
      },
      {
        id: 'same_unit_assumption',
        label: 'Amount units',
        status: amounts === null ? 'passed' : 'warning',
        detail: computed.currency.assumption,
      },
    ],
    notes: [
      'Matched and exception sheets contain one record per source row, not one per pair. Pair counts and signed-delta totals count each pair once.',
      `String key normalization: ${computed.args.keyNormalization}. Numeric, boolean, and date types never coerce to strings; date keys retain milliseconds.`,
      `Absolute same-unit tolerance: ${computed.args.tolerance}. Signed left-minus-right deltas are retained even for tolerance matches.`,
      'Live supported deterministic formulas are evaluated against the original workbook; cached values and formula text are not copied into the report.',
      'This is a deterministic reconciliation of the supplied bindings, not a finance opinion, fuzzy match, FX conversion, duplicate aggregation, or settlement instruction.',
    ],
  };
}

function summaryValues(
  computed: ComputedReconciliation,
  artifact: WorkbookArtifact,
): CellValue[][] {
  const rows: CellValue[][] = [['Section', 'Metric', 'Value', 'Detail']];
  const add = (section: string, metric: string, value: CellValue, detail = '') =>
    rows.push([section, metric, value, detail]);
  add(
    'Report',
    'Title',
    artifact.title,
    'Finished value-only report; source sheets are unchanged.',
  );
  for (const [index, side] of (['left', 'right'] as const).entries()) {
    add('Bindings', `${side} full source range`, qualifiedRange(computed.sources[index]!));
    add(
      'Bindings',
      `${side} header row`,
      computed.args[`${side}HeaderRow`],
      'Every allocated source row after this header is included.',
    );
    add('Bindings', `${side} key columns`, computed.args[`${side}Keys`].join(', '));
    add('Bindings', `${side} amount column`, computed.args[`${side}Amount`] ?? 'Not configured');
    add(
      'Bindings',
      `${side} currency header indication`,
      computed.currency[side] ?? 'Not established',
      'Header labels alone cannot certify every amount unit.',
    );
    add('Counts', `${side} source rows`, computed.counts[side].sourceRows);
    add('Counts', `${side} accounted rows`, computed.counts[side].accountedRows);
    for (const outcome of RECONCILIATION_OUTCOMES)
      add('Counts', `${side} ${outcome}`, computed.counts[side].outcomes[outcome]);
  }
  add('Bindings', 'Key normalization', computed.args.keyNormalization);
  add(
    'Bindings',
    'Absolute same-unit tolerance',
    computed.args.tolerance,
    'Tolerance is not a percentage and does not alter raw deltas.',
  );
  add('Counts', 'Matched pairs', computed.counts.pairs.matched);
  add('Counts', 'Amount-mismatch pairs', computed.counts.pairs.amount_mismatch);
  add('Counts', 'Invalid-amount pairs', computed.counts.pairs.invalid_amount);
  add('Counts', 'Duplicate groups', computed.counts.duplicateGroups);
  add('Counts', 'Exception source rows', computed.counts.exceptionRows);
  if (computed.amounts) {
    for (const side of ['left', 'right'] as const) {
      const control = computed.amounts[side];
      for (const population of ['total', 'matched', 'exceptions'] as const) {
        const total = control[population];
        add(
          'Amounts',
          `${side} ${population} signed total`,
          total.value,
          total.overflow
            ? 'Unavailable: finite-range overflow, not zero.'
            : 'Sum of numeric inclusions only.',
        );
        add('Amounts', `${side} ${population} numeric count`, total.numericCount);
        add('Amounts', `${side} ${population} excluded count`, total.excludedCount);
        add(
          'Amounts',
          `${side} ${population} numeric-text conversions`,
          total.numericTextConversions,
        );
        add(
          'Amounts',
          `${side} ${population} absolute sum`,
          total.absoluteSum,
          total.absoluteSum === null
            ? 'Unavailable: finite-range overflow.'
            : 'For numerical controls, not a financial conclusion.',
        );
        add(
          'Amounts',
          `${side} ${population} roundoff bound`,
          total.roundoffBound,
          total.roundoffBound === null
            ? 'Unavailable: finite-range overflow.'
            : 'Absolute numerical accumulation error bound; separate from business tolerance.',
        );
      }
      add(
        'Proof',
        `${side} population counts balanced`,
        control.proof.countsBalanced &&
          control.proof.numericCountsBalanced &&
          control.proof.exclusionsBalanced,
      );
      add(
        'Proof',
        `${side} total - matched - exceptions`,
        control.proof.residual,
        control.proof.residual === null
          ? 'Unavailable: overflow in a required total.'
          : 'Signed control-total residual.',
      );
      add('Proof', `${side} partition roundoff bound`, control.proof.roundoffBound);
      add(
        'Proof',
        `${side} control-total proof`,
        control.proof.totalsBalanced === null ? 'Unavailable' : control.proof.totalsBalanced,
      );
    }
    for (const population of ['pairedDeltas', 'matchedDeltas'] as const) {
      const total = computed.amounts[population];
      add(
        'Deltas',
        `${population} signed sum`,
        total.value,
        'Each pair counted once; left minus right, including tolerance matches.',
      );
      add('Deltas', `${population} finite count`, total.numericCount);
      add(
        'Deltas',
        `${population} excluded count`,
        total.excludedCount,
        'Invalid-amount or overflowing pair differences, not zero.',
      );
      add('Deltas', `${population} roundoff bound`, total.roundoffBound);
    }
  } else add('Amounts', 'Mode', 'Key-only; amounts not configured');
  for (const check of artifact.checks) add('Checks', check.label, check.status, check.detail);
  for (const warning of computed.warnings)
    add('Warnings', warning.code, 'warning', warning.message);
  return rows;
}

function methodologyValues(workbook: Workbook, computed: ComputedReconciliation): CellValue[][] {
  const args = computed.args;
  return [
    ['Topic', 'Method', 'Binding / limit'],
    [
      'Purpose',
      'Deterministic one-to-one reconciliation of two explicitly bound source sheets.',
      'No financial conclusion or FX conversion.',
    ],
    [
      'Left binding',
      `${args.leftSheet}; header row ${args.leftHeaderRow}; keys ${args.leftKeys.join(', ')}; amount ${args.leftAmount ?? 'not configured'}`,
      qualifiedRange(computed.sources[0]!),
    ],
    [
      'Right binding',
      `${args.rightSheet}; header row ${args.rightHeaderRow}; keys ${args.rightKeys.join(', ')}; amount ${args.rightAmount ?? 'not configured'}`,
      qualifiedRange(computed.sources[1]!),
    ],
    [
      'Data population',
      'All allocated source rows after each header, including blank and ragged rows; title/header rows are not records.',
      'One report row per source row, either matched or exception.',
    ],
    [
      'Key encoding',
      'Length-framed typed scalars and length-framed components; no delimiter collisions, numeric/string coercion, or loss of leading zeros.',
      'Date keys retain epoch milliseconds; booleans and signed zero retain their scalar identity.',
    ],
    [
      'String normalization',
      'Exact is the default. Trim removes leading/trailing whitespace; trim_casefold additionally applies deterministic JavaScript Unicode lowercase (not locale-specific or full Unicode folding). Other scalar types are unchanged.',
      args.keyNormalization,
    ],
    [
      'Blank/invalid keys',
      'Any blank component, non-finite number, invalid date, formula error, unsupported/volatile formula or dependent makes the row invalid_key.',
      'Whitespace-only keys are invalid, even in exact mode.',
    ],
    [
      'Duplicate keys',
      'If either side has more than one row for a key, every row on both sides is duplicate_key.',
      'No first-wins, automatic aggregation, ambiguous pairing, or fuzzy matching.',
    ],
    [
      'Amount parsing',
      'Finite numbers or conservative numeric text via the engine toNumericOrNull rule (including valid thousands groups and exponent notation). Negatives and zero preserved.',
      'Blank/error/boolean/date/unsupported/volatile/currency-formatted text is excluded, not zero; conversions counted.',
    ],
    [
      'Units',
      computed.currency.assumption,
      `Explicit currency header indications: left ${computed.currency.left ?? 'unknown'}, right ${computed.currency.right ?? 'unknown'}. Conflicting common currencies block the entire job.`,
    ],
    [
      'Tolerance',
      'A unique pair matches when both amounts are numeric and absolute(left - right) <= tolerance. Signed raw delta is retained on both source records.',
      args.tolerance,
    ],
    [
      'Outcome precedence',
      'Invalid key, then duplicate group, then unpaired key; only paired unique keys undergo amount comparison.',
      'Unpaired rows with invalid amounts remain left_only/right_only, with their amount exclusion reason visible.',
    ],
    [
      'Control totals',
      'Scale-first Neumaier compensated signed sums. Numeric/excluded counts, conversions, matched/exception populations, absolute sums, and roundoff bounds are shown.',
      'Unrepresentable totals/differences are unavailable with warnings; no infinity, silent zero, or rounded-away delta.',
    ],
    [
      'Partition proof',
      'Total - matched - exceptions must be within the sum of their numerical roundoff bounds plus residual-accumulation roundoff.',
      'Unavailable numerical proof is a warning, not a pass. Business discrepancies are separate warnings, not broken row accounting.',
    ],
    [
      'Formulas',
      'Read live deterministic supported scalar values against the original snapshot. Unsupported and clock/random formulas and dependents are excluded, including masked dependents.',
      'Formula subset and dependency-analysis limits are those of the engine deterministic value reader; never use caches or copy formula text.',
    ],
    [
      'Date epoch',
      'Source workbook metadata and all original cells/formulas remain unchanged; generated sheets contain values only.',
      workbook.dateSystem ?? '1900 (default)',
    ],
    [
      'Output ordering',
      'Left source rows ascending, then right source rows ascending, partitioned into matched and exceptions without reordering source sheets.',
      'Matched pair counts are half the matched source-record count. Delta totals count each pair once.',
    ],
    [
      'Source references',
      'Each record includes side, source sheet/row/full-row range, key-cell references, optional amount-cell reference and unique counterpart reference.',
      'References are inert text, not formulas or hyperlinks.',
    ],
    [
      'Atomic delivery',
      'Exactly four uniquely named sheets appended in summary/matched/exceptions/methodology order, with reversible workbook snapshot patch and preservation invariant.',
      'No partial output, source rewrite, or silent truncation.',
    ],
    [
      'Capacity',
      'Projected original plus generated allocated model cells must fit the guard; explicit blank cells count. Generated rectangular report dimensions are checked before allocation.',
      RECONCILIATION_LIMITS.allocatedCells,
    ],
    [
      'Excel limits',
      'Generated names <=31 characters and valid; per-sheet row/column limits and per-output-cell text length enforced.',
      `${RECONCILIATION_LIMITS.sheetRows} rows; ${RECONCILIATION_LIMITS.sheetColumns} columns; ${RECONCILIATION_LIMITS.cellTextLength} text characters.`,
    ],
    [
      'Report schemas',
      'Summary: Section/Metric/Value/Detail. Methodology: Topic/Method/Binding or limit. Matched and exceptions: identical source-record schemas.',
      RECORD_HEADERS.join('; '),
    ],
  ];
}

function assertOutputValues(values: CellValue[]): void {
  for (const value of values) {
    if (typeof value === 'string' && value.length > RECONCILIATION_LIMITS.cellTextLength)
      throw new Error(
        'A reconciliation output cell exceeds the Excel 32,767-character text limit. No report was appended; no value was truncated.',
      );
    if (typeof value === 'number' && !Number.isFinite(value))
      throw new Error('A reconciliation output value is not finite. No report was appended.');
  }
}

/** Compute dimensions and validate the whole job before allocating any generated Cell objects. */
function prepareReport(workbook: Workbook, args: ReconcileSheetsArgs): ReportPlan {
  const computed = computeReconciliation(workbook, args);
  const names = reportNames(workbook, computed.args.reportPrefix);
  const artifact = artifactFor(computed, names);
  if (artifact.checks.some((check) => check.status === 'failed'))
    throw new Error('Reconciliation accounting proof failed. No report was appended.');
  const summary = summaryValues(computed, artifact);
  const methodology = methodologyValues(workbook, computed);
  const matchedCount =
    computed.counts.left.outcomes.matched + computed.counts.right.outcomes.matched;
  const sizes = {
    summary: { rows: summary.length, columns: 4 },
    matched: { rows: matchedCount + 1, columns: RECORD_HEADERS.length },
    exceptions: { rows: computed.counts.exceptionRows + 1, columns: RECORD_HEADERS.length },
    methodology: { rows: methodology.length, columns: 3 },
  };
  let generatedCells = 0;
  const ranges = ROLES.map((role) => {
    const size = sizes[role];
    if (
      size.rows > RECONCILIATION_LIMITS.sheetRows ||
      size.columns > RECONCILIATION_LIMITS.sheetColumns
    )
      throw new Error(
        `Reconciliation ${role} sheet exceeds Excel worksheet limits. No report was appended.`,
      );
    generatedCells += size.rows * size.columns;
    return {
      sheet: names[role],
      startRow: 1,
      endRow: size.rows,
      startColumn: 'A',
      endColumn: indexToColumn(size.columns - 1),
    };
  });
  const projected = reconciliationAllocatedCells(workbook) + generatedCells;
  if (projected > RECONCILIATION_LIMITS.allocatedCells)
    throw new Error(
      `Reconciliation would allocate ${projected.toLocaleString('en-US')} workbook cells, above the ${RECONCILIATION_LIMITS.allocatedCells.toLocaleString('en-US')} cell limit. No report was appended or truncated.`,
    );
  for (const values of [...summary, ...methodology]) assertOutputValues(values);
  for (const row of computed.rows) assertOutputValues(recordValues(row));
  return { computed, names, summary, methodology, ranges, generatedCells, artifact };
}

function validationFailure(error: unknown): ValidationResult {
  return {
    valid: false,
    errors: [
      issue(
        'reconciliation-job',
        error instanceof Error ? error.message : 'Reconciliation failed.',
      ),
    ],
    warnings: [],
  };
}

function valueCells(values: CellValue[]): Cell[] {
  return values.map((value) => createCell(value));
}
function generatedSheets(plan: ReportPlan): Sheet[] {
  const matched = [valueCells(RECORD_HEADERS)];
  const exceptions = [valueCells(RECORD_HEADERS)];
  for (const row of plan.computed.rows)
    (row.outcome === 'matched' ? matched : exceptions).push(valueCells(recordValues(row)));
  return [
    { name: plan.names.summary, rows: plan.summary.map(valueCells) },
    { name: plan.names.matched, rows: matched },
    { name: plan.names.exceptions, rows: exceptions },
    { name: plan.names.methodology, rows: plan.methodology.map(valueCells) },
  ];
}

function preservedSources(before: Workbook, after: Workbook): string[] {
  const errors: string[] = [];
  if (before.dateSystem !== after.dateSystem)
    errors.push('Original workbook date epoch metadata was not preserved exactly.');
  for (const [sheetIndex, sheet] of before.sheets.entries()) {
    const actual = after.sheets[sheetIndex];
    if (!actual || actual.name !== sheet.name || actual.rows.length !== sheet.rows.length) {
      errors.push(`Original source sheet order/shape was not preserved: "${sheet.name}".`);
      continue;
    }
    for (const [rowIndex, row] of sheet.rows.entries()) {
      const actualRow = actual.rows[rowIndex];
      if (!actualRow || actualRow.length !== row.length) {
        errors.push(`Original row shape changed: "${sheet.name}" row ${rowIndex + 1}.`);
        continue;
      }
      for (const [column, cell] of row.entries()) {
        const actualCell = actualRow[column];
        const sameValue =
          actualCell !== undefined &&
          (cellValueEquals(cell.value, actualCell.value) ||
            (cell.value instanceof Date &&
              actualCell.value instanceof Date &&
              Number.isNaN(cell.value.getTime()) &&
              Number.isNaN(actualCell.value.getTime())));
        if (
          !actualCell ||
          !sameValue ||
          cell.type !== actualCell.type ||
          cell.formula !== actualCell.formula ||
          cell.numberFormat !== actualCell.numberFormat
        ) {
          errors.push(
            `Original cell changed: "${sheet.name}" ${indexToColumn(column)}${rowIndex + 1}.`,
          );
          if (errors.length >= 20) return errors;
        }
      }
    }
  }
  return errors;
}

export const reconcileSheetsOperation: Operation<ReconcileSheetsArgs> = {
  name: 'reconcile_sheets',
  schema: reconcileSheetsArgsSchema,
  targetRanges(workbook, args) {
    return prepareReport(workbook, args).ranges;
  },
  validate(workbook, args) {
    try {
      const plan = prepareReport(workbook, args);
      return { valid: true, errors: [], warnings: plan.computed.warnings };
    } catch (error) {
      return validationFailure(error);
    }
  },
  preview(workbook, args) {
    try {
      const plan = prepareReport(workbook, args);
      // Only materialize the bounded preview sample; the complete projected cell count is exact.
      const sample: Workbook = {
        ...workbook,
        sheets: [
          ...workbook.sheets,
          { name: plan.names.summary, rows: plan.summary.slice(0, 5).map(valueCells) },
        ],
      };
      return {
        ...previewForTransition(workbook, sample, plan.ranges, plan.computed.warnings),
        affectedCells: plan.generatedCells,
        requiresConfirmation: plan.generatedCells > 200,
      };
    } catch (error) {
      return previewForTransition(workbook, workbook, [], [], validationFailure(error).errors);
    }
  },
  apply(workbook, args) {
    const plan = prepareReport(workbook, args);
    const after: Workbook = { ...workbook, sheets: [...workbook.sheets, ...generatedSheets(plan)] };
    return transitionResult(workbook, after, {
      artifacts: [plan.artifact],
      affectedCells: plan.generatedCells,
      skippedCells: 0,
      unchangedCells: reconciliationAllocatedCells(workbook),
      warnings: plan.computed.warnings,
      addedRows: plan.ranges.reduce((sum, range) => sum + range.endRow, 0),
      matchedRows:
        plan.computed.counts.left.outcomes.matched + plan.computed.counts.right.outcomes.matched,
      unmatchedRows: plan.computed.counts.exceptionRows,
    });
  },
  invariants(before, after, args): InvariantResult {
    try {
      const plan = prepareReport(before, args);
      const errors = preservedSources(before, after);
      if (after.sheets.length !== before.sheets.length + 4)
        errors.push('Reconciliation must append exactly four report sheets.');
      for (const [index, range] of plan.ranges.entries()) {
        const sheet = after.sheets[before.sheets.length + index];
        if (!sheet || sheet.name !== range.sheet || sheet.rows.length !== range.endRow) {
          errors.push(
            `Generated report sheet name/order/row accounting changed for "${range.sheet}".`,
          );
          continue;
        }
        const expectedWidth = index === 0 ? 4 : index === 3 ? 3 : RECORD_HEADERS.length;
        if (
          sheet.rows.some(
            (row) =>
              row.length !== expectedWidth ||
              row.some((cell) => cell.formula !== undefined || cell.type === 'formula'),
          )
        )
          errors.push(
            `Generated sheet "${sheet.name}" must contain the complete rectangular value-only report.`,
          );
      }
      const ordinary = runInvariants(before, after, { targetRanges: plan.ranges });
      errors.push(...ordinary.errors);
      return { valid: errors.length === 0, errors };
    } catch (error) {
      return {
        valid: false,
        errors: [error instanceof Error ? error.message : 'Reconciliation invariant failed.'],
      };
    }
  },
};
