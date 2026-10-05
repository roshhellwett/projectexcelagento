import { z } from 'zod';

import { createDeterministicWorkbookValueReader } from './formula/deterministic-value-reader.js';
import { toNumericOrNull } from './formula/functions.js';
import {
  fullSheetRange,
  headerRowError,
  issue,
  validateColumn,
  validateSheet,
} from './operation-utils.js';
import type {
  CellLocation,
  CellRange,
  CellValue,
  ValidationIssue,
  ValidationResult,
  Workbook,
} from './types.js';
import { columnToIndex, getSheet, maxColumnCount } from './workbook.js';

export const RECONCILIATION_LIMITS = Object.freeze({
  allocatedCells: 1_500_000,
  sheetRows: 1_048_576,
  sheetColumns: 16_384,
  cellTextLength: 32_767,
});

const columnSchema = z
  .string()
  .trim()
  .regex(/^[A-Za-z]+$/, 'Columns must be letters')
  .transform((value) => value.toUpperCase())
  .refine(
    (value) => (columnToIndex(value) ?? Infinity) < RECONCILIATION_LIMITS.sheetColumns,
    'Column exceeds Excel column XFD',
  );

export const reconcileSheetsArgsSchema = z
  .object({
    leftSheet: z.string().min(1),
    rightSheet: z.string().min(1),
    leftKeys: z.array(columnSchema).min(1).max(5),
    rightKeys: z.array(columnSchema).min(1).max(5),
    leftAmount: columnSchema.optional(),
    rightAmount: columnSchema.optional(),
    leftHeaderRow: z.number().int().positive().max(RECONCILIATION_LIMITS.sheetRows).default(1),
    rightHeaderRow: z.number().int().positive().max(RECONCILIATION_LIMITS.sheetRows).default(1),
    tolerance: z.number().finite().nonnegative().default(0),
    keyNormalization: z.enum(['exact', 'trim', 'trim_casefold']).default('exact'),
    reportPrefix: z.string().trim().min(1).default('Reconciliation'),
  })
  .superRefine((args, context) => {
    if (args.leftKeys.length !== args.rightKeys.length) {
      context.addIssue({
        code: 'custom',
        path: ['rightKeys'],
        message: 'Left and right key mappings must have equal length.',
      });
    }
    if ((args.leftAmount === undefined) !== (args.rightAmount === undefined)) {
      context.addIssue({
        code: 'custom',
        path: ['rightAmount'],
        message: 'Amount columns must be provided together.',
      });
    }
    if (args.leftSheet.toLowerCase() === args.rightSheet.toLowerCase()) {
      context.addIssue({
        code: 'custom',
        path: ['rightSheet'],
        message: 'Reconciliation requires two different sheets.',
      });
    }
    for (const side of ['left', 'right'] as const) {
      const keys = args[`${side}Keys`];
      if (
        new Set(keys).size !== keys.length ||
        (args[`${side}Amount`] !== undefined && keys.includes(args[`${side}Amount`]!))
      ) {
        context.addIssue({
          code: 'custom',
          path: [`${side}Keys`],
          message: 'Key columns must be distinct and may not also be the amount column.',
        });
      }
    }
  });
export type ReconcileSheetsArgs = z.infer<typeof reconcileSheetsArgsSchema>;

export type ReconciliationSide = 'left' | 'right';
export const RECONCILIATION_OUTCOMES = [
  'matched',
  'amount_mismatch',
  'invalid_amount',
  'left_only',
  'right_only',
  'duplicate_key',
  'invalid_key',
] as const;
export type ReconciliationOutcome = (typeof RECONCILIATION_OUTCOMES)[number];
export type ValueExclusion =
  'blank' | 'non_numeric' | 'non_finite' | 'invalid_date' | 'unsupported' | 'volatile' | 'error';

export interface ReconciliationAmount {
  /** Null means excluded, never an implicit zero. */
  value: number | null;
  reason: ValueExclusion | null;
  numericTextConverted: boolean;
}

export interface ReconciliationRow {
  side: ReconciliationSide;
  sheet: string;
  sourceRow: number;
  sourceRange: CellRange;
  keyCells: CellLocation[];
  /** Live original scalar values, not cached formula values or formula text. */
  keyValues: CellValue[];
  /** A typed, length-framed composite; null if any component is invalid. */
  key: string | null;
  keyProblems: { column: string; reason: ValueExclusion }[];
  amountCell: CellLocation | null;
  amount: ReconciliationAmount | null;
  outcome: ReconciliationOutcome;
  counterpart: { sheet: string; sourceRow: number; sourceRange: CellRange } | null;
  /** Always left minus right, even on right-side records and tolerance matches. */
  delta: number | null;
  deltaReason: 'not_configured' | 'not_paired' | 'invalid_amount' | 'overflow' | null;
  detail: string;
}

export interface ReconciliationNumericTotal {
  value: number | null;
  numericCount: number;
  excludedCount: number;
  numericTextConversions: number;
  absoluteSum: number | null;
  /** Conservative absolute bound for binary floating-point accumulation, not business tolerance. */
  roundoffBound: number | null;
  overflow: boolean;
}

export interface ReconciliationAmountControls {
  total: ReconciliationNumericTotal;
  matched: ReconciliationNumericTotal;
  exceptions: ReconciliationNumericTotal;
  proof: {
    countsBalanced: boolean;
    numericCountsBalanced: boolean;
    exclusionsBalanced: boolean;
    residual: number | null;
    roundoffBound: number | null;
    totalsBalanced: boolean | null;
  };
}

export interface ReconciliationSideCounts {
  sourceRows: number;
  outcomes: Record<ReconciliationOutcome, number>;
  accountedRows: number;
}

export interface ComputedReconciliation {
  args: ReconcileSheetsArgs;
  /** Full allocated source sheets, including headers/title rows. Data rows begin after the bound header. */
  sources: CellRange[];
  /** Stable order: all left source rows ascending, then all right source rows ascending. */
  rows: ReconciliationRow[];
  counts: {
    left: ReconciliationSideCounts;
    right: ReconciliationSideCounts;
    pairs: { matched: number; amount_mismatch: number; invalid_amount: number };
    duplicateGroups: number;
    exceptionRows: number;
    rowAccountingComplete: boolean;
  };
  amounts: {
    left: ReconciliationAmountControls;
    right: ReconciliationAmountControls;
    /** Each pair is counted once, never twice for its two source records. */
    pairedDeltas: ReconciliationNumericTotal;
    matchedDeltas: ReconciliationNumericTotal;
  } | null;
  currency: { left: string | null; right: string | null; assumption: string };
  warnings: ValidationIssue[];
}

/** Allocated model cells, including explicit blanks and ragged row widths (not rectangular used-range padding). */
export function reconciliationAllocatedCells(workbook: Workbook): number {
  let total = 0;
  for (const sheet of workbook.sheets) for (const row of sheet.rows) total += row.length;
  return total;
}

export function validateReconciliationInputs(
  workbook: Workbook,
  input: ReconcileSheetsArgs,
): ValidationResult {
  const parsed = reconcileSheetsArgsSchema.safeParse(input);
  if (!parsed.success)
    return {
      valid: false,
      errors: parsed.error.issues.map((problem) => issue('reconciliation-schema', problem.message)),
      warnings: [],
    };
  const args = parsed.data;
  const errors: ValidationIssue[] = [];
  const names = new Set<string>();
  for (const sheet of workbook.sheets) {
    if (names.has(sheet.name.toLowerCase()))
      errors.push(
        issue('ambiguous-sheet', `Sheet names are ambiguous ignoring case: "${sheet.name}".`),
      );
    names.add(sheet.name.toLowerCase());
    if (
      sheet.rows.length > RECONCILIATION_LIMITS.sheetRows ||
      maxColumnCount(sheet.rows) > RECONCILIATION_LIMITS.sheetColumns
    ) {
      errors.push(
        issue('excel-sheet-limit', `Sheet "${sheet.name}" exceeds Excel row or column limits.`),
      );
    }
  }
  if (reconciliationAllocatedCells(workbook) > RECONCILIATION_LIMITS.allocatedCells) {
    errors.push(
      issue(
        'reconciliation-cell-limit',
        `Workbook already exceeds the ${RECONCILIATION_LIMITS.allocatedCells.toLocaleString('en-US')} allocated-cell limit.`,
      ),
    );
  }
  for (const side of ['left', 'right'] as const) {
    const sheet = args[`${side}Sheet`];
    errors.push(...validateSheet(workbook, sheet));
    if (!getSheet(workbook, sheet)) continue;
    errors.push(...headerRowError(workbook, sheet, args[`${side}HeaderRow`]));
    for (const column of [
      ...args[`${side}Keys`],
      ...(args[`${side}Amount`] ? [args[`${side}Amount`]!] : []),
    ]) {
      errors.push(...validateColumn(workbook, sheet, column));
    }
  }
  return { valid: errors.length === 0, errors, warnings: [] };
}

// Codes and unambiguous currency symbols only: a bare $ or ¥ cannot establish a unit.
const COMMON_CURRENCIES = new Set(
  'USD EUR GBP JPY CNY AUD CAD CHF INR NZD HKD SGD KRW RUB BRL MXN ZAR AED SAR SEK NOK DKK PLN TRY IDR THB MYR PHP VND TWD ILS'.split(
    ' ',
  ),
);
function currenciesInHeader(value: CellValue): string[] {
  if (typeof value !== 'string') return [];
  const upper = value.toUpperCase();
  const codes = new Set<string>();
  for (const token of upper.match(/[A-Z]+/g) ?? []) {
    if (COMMON_CURRENCIES.has(token)) codes.add(token);
    if (token === 'RMB') codes.add('CNY');
  }
  for (const [pattern, currency] of [
    [/US\$/, 'USD'],
    [/C\$/, 'CAD'],
    [/A\$/, 'AUD'],
    [/NZ\$/, 'NZD'],
    [/HK\$/, 'HKD'],
    [/S\$/, 'SGD'],
    [/€/, 'EUR'],
    [/£/, 'GBP'],
    [/₹/, 'INR'],
    [/₩/, 'KRW'],
    [/₽/, 'RUB'],
    [/₺/, 'TRY'],
  ] as const) {
    if (pattern.test(upper)) codes.add(currency);
  }
  return [...codes].sort();
}

function cloneValue(value: CellValue): CellValue {
  return value instanceof Date ? new Date(value.getTime()) : value;
}

function keyComponent(
  value: CellValue,
  mode: ReconcileSheetsArgs['keyNormalization'],
): { framed: string; reason: ValueExclusion | null } {
  let type: string;
  let text: string;
  if (value === null || (typeof value === 'string' && value.trim() === ''))
    return { framed: '', reason: 'blank' };
  if (value instanceof Date) {
    if (!Number.isFinite(value.getTime())) return { framed: '', reason: 'invalid_date' };
    type = 'date';
    text = String(value.getTime());
  } else if (typeof value === 'number') {
    if (!Number.isFinite(value)) return { framed: '', reason: 'non_finite' };
    type = 'number';
    text = Object.is(value, -0) ? '-0' : String(value);
  } else if (typeof value === 'boolean') {
    type = 'boolean';
    text = String(value);
  } else if (typeof value === 'string') {
    type = 'string';
    text = mode === 'exact' ? value : value.trim();
    if (mode === 'trim_casefold') text = text.toLowerCase();
  } else return { framed: '', reason: 'unsupported' };
  return { framed: `${type.length}:${type}${text.length}:${text}`, reason: null };
}

function amountValue(
  value: CellValue,
  reason: 'unsupported' | 'volatile' | 'error' | null,
): ReconciliationAmount {
  if (reason) return { value: null, reason, numericTextConverted: false };
  if (value === null || (typeof value === 'string' && value.trim() === ''))
    return { value: null, reason: 'blank', numericTextConverted: false };
  if (typeof value === 'number' && !Number.isFinite(value))
    return { value: null, reason: 'non_finite', numericTextConverted: false };
  const numeric = toNumericOrNull(value);
  return numeric === null || !Number.isFinite(numeric)
    ? { value: null, reason: 'non_numeric', numericTextConverted: false }
    : { value: numeric, reason: null, numericTextConverted: typeof value === 'string' };
}

/**
 * Scale-first Neumaier compensated sum. Intermediate additions remain bounded by the row count,
 * even for MAX_VALUE + MAX_VALUE - MAX_VALUE. Overflow is explicit only at conversion back to
 * source units; neither an infinity nor a fabricated zero ever enters the report.
 */
class CompensatedTotal {
  private scale = 0;
  private sum = 0;
  private compensation = 0;
  private absolute = 0;
  private count = 0;
  private excluded = 0;
  private conversions = 0;

  add(value: number | null, converted = false): void {
    if (value === null) {
      this.excluded += 1;
      return;
    }
    this.count += 1;
    if (converted) this.conversions += 1;
    const magnitude = Math.abs(value);
    if (magnitude > this.scale) {
      const ratio = this.scale / magnitude;
      this.sum *= ratio;
      this.compensation *= ratio;
      this.absolute *= ratio;
      this.scale = magnitude;
    }
    if (this.scale === 0) return;
    const scaled = value / this.scale;
    const next = this.sum + scaled;
    this.compensation +=
      Math.abs(this.sum) >= Math.abs(scaled) ? this.sum - next + scaled : scaled - next + this.sum;
    this.sum = next;
    this.absolute += Math.abs(scaled);
  }

  result(): ReconciliationNumericTotal {
    const raw = (this.sum + this.compensation) * this.scale;
    const absoluteSum = this.absolute * this.scale;
    const bound =
      (8 * this.count + 16) * Number.EPSILON * this.absolute * this.scale +
      this.count * Number.MIN_VALUE;
    return {
      value: Number.isFinite(raw) ? raw : null,
      numericCount: this.count,
      excludedCount: this.excluded,
      numericTextConversions: this.conversions,
      absoluteSum: Number.isFinite(absoluteSum) ? absoluteSum : null,
      roundoffBound: Number.isFinite(bound) ? bound : null,
      overflow: !Number.isFinite(raw),
    };
  }
}

function amountControls(
  rows: ReconciliationRow[],
  side: ReconciliationSide,
): ReconciliationAmountControls {
  const all = new CompensatedTotal();
  const matched = new CompensatedTotal();
  const exceptions = new CompensatedTotal();
  let population = 0;
  for (const row of rows) {
    if (row.side !== side) continue;
    population += 1;
    const amount = row.amount!;
    all.add(amount.value, amount.numericTextConverted);
    (row.outcome === 'matched' ? matched : exceptions).add(
      amount.value,
      amount.numericTextConverted,
    );
  }
  const total = all.result();
  const matching = matched.result();
  const except = exceptions.result();
  const residualSum = new CompensatedTotal();
  const computable = total.value !== null && matching.value !== null && except.value !== null;
  if (computable) {
    residualSum.add(total.value);
    residualSum.add(-matching.value!);
    residualSum.add(-except.value!);
  }
  const residual = computable ? residualSum.result().value : null;
  const bounds = [
    total.roundoffBound,
    matching.roundoffBound,
    except.roundoffBound,
    residualSum.result().roundoffBound,
  ];
  const rawBound = bounds.every((bound) => bound !== null)
    ? bounds.reduce<number>((sum, bound) => sum + bound!, 0)
    : Infinity;
  const roundoffBound = Number.isFinite(rawBound) ? rawBound : null;
  return {
    total,
    matched: matching,
    exceptions: except,
    proof: {
      countsBalanced: total.numericCount + total.excludedCount === population,
      numericCountsBalanced: total.numericCount === matching.numericCount + except.numericCount,
      exclusionsBalanced: total.excludedCount === matching.excludedCount + except.excludedCount,
      residual,
      roundoffBound,
      totalsBalanced:
        residual === null || roundoffBound === null ? null : Math.abs(residual) <= roundoffBound,
    },
  };
}

/** Computes every source-row outcome without changing a cell or relying on a formula cache. */
export function computeReconciliation(
  workbook: Workbook,
  input: ReconcileSheetsArgs,
): ComputedReconciliation {
  const args = reconcileSheetsArgsSchema.parse(input);
  const validation = validateReconciliationInputs(workbook, args);
  if (!validation.valid)
    throw new Error(validation.errors.map((problem) => problem.message).join(' '));
  const read = createDeterministicWorkbookValueReader(workbook);
  const hasAmounts = args.leftAmount !== undefined;
  const currency = {
    left: null as string | null,
    right: null as string | null,
    assumption: hasAmounts
      ? 'Amounts are compared in the same units as supplied. No currency inference, FX conversion, or financial conclusion is performed.'
      : 'Key-only reconciliation; no amount or currency comparison is performed.',
  };
  if (hasAmounts) {
    for (const side of ['left', 'right'] as const) {
      const header = read(
        args[`${side}Sheet`],
        columnToIndex(args[`${side}Amount`]!)!,
        args[`${side}HeaderRow`],
      );
      const codes = header.reason === null ? currenciesInHeader(header.value) : [];
      if (codes.length > 1)
        throw new Error(
          `The ${side} amount header explicitly indicates multiple currencies (${codes.join(', ')}); same-unit reconciliation is unsafe.`,
        );
      currency[side] = codes[0] ?? null;
    }
    if (currency.left !== null && currency.right !== null && currency.left !== currency.right) {
      throw new Error(
        `Amount headers explicitly indicate different currencies (${currency.left} versus ${currency.right}). Reconciliation requires same-unit amounts; no FX conversion is performed.`,
      );
    }
  }

  const rows: ReconciliationRow[] = [];
  const groups = new Map<string, { left: ReconciliationRow[]; right: ReconciliationRow[] }>();
  const sources = [
    fullSheetRange(workbook, args.leftSheet),
    fullSheetRange(workbook, args.rightSheet),
  ];
  for (const [sideIndex, side] of (['left', 'right'] as const).entries()) {
    const sheet = getSheet(workbook, args[`${side}Sheet`])!;
    const columns = args[`${side}Keys`];
    const indices = columns.map((column) => columnToIndex(column)!);
    const amountColumn = args[`${side}Amount`];
    const amountIndex = amountColumn === undefined ? undefined : columnToIndex(amountColumn)!;
    for (
      let sourceRow = args[`${side}HeaderRow`] + 1;
      sourceRow <= sheet.rows.length;
      sourceRow += 1
    ) {
      const keyValues: CellValue[] = [];
      const keyProblems: ReconciliationRow['keyProblems'] = [];
      const frames: string[] = [];
      for (const [index, columnIndex] of indices.entries()) {
        const current = read(sheet.name, columnIndex, sourceRow);
        keyValues.push(cloneValue(current.value));
        const component = keyComponent(current.value, args.keyNormalization);
        const reason = current.reason ?? component.reason;
        if (reason) keyProblems.push({ column: columns[index]!, reason });
        frames.push(component.framed);
      }
      const key =
        keyProblems.length > 0 ? null : frames.map((frame) => `${frame.length}:${frame}`).join('');
      const currentAmount =
        amountIndex === undefined ? null : read(sheet.name, amountIndex, sourceRow);
      const sourceRange = { ...sources[sideIndex]!, startRow: sourceRow, endRow: sourceRow };
      const row: ReconciliationRow = {
        side,
        sheet: sheet.name,
        sourceRow,
        sourceRange,
        keyCells: columns.map((column) => ({ sheet: sheet.name, row: sourceRow, column })),
        keyValues,
        key,
        keyProblems,
        amountCell:
          amountColumn === undefined
            ? null
            : { sheet: sheet.name, row: sourceRow, column: amountColumn },
        amount:
          currentAmount === null ? null : amountValue(currentAmount.value, currentAmount.reason),
        outcome: key === null ? 'invalid_key' : side === 'left' ? 'left_only' : 'right_only',
        counterpart: null,
        delta: null,
        deltaReason: hasAmounts ? 'not_paired' : 'not_configured',
        detail:
          key === null
            ? `Invalid key: ${keyProblems.map((problem) => `${problem.column}: ${problem.reason}`).join('; ')}.`
            : `Valid key has no unique ${side === 'left' ? 'right' : 'left'} counterpart.`,
      };
      rows.push(row);
      if (key !== null) {
        const group = groups.get(key) ?? { left: [], right: [] };
        group[side].push(row);
        groups.set(key, group);
      }
    }
  }

  let duplicateGroups = 0;
  const pairs = { matched: 0, amount_mismatch: 0, invalid_amount: 0 };
  const pairedDeltas = new CompensatedTotal();
  const matchedDeltas = new CompensatedTotal();
  for (const group of groups.values()) {
    if (group.left.length > 1 || group.right.length > 1) {
      duplicateGroups += 1;
      for (const row of [...group.left, ...group.right]) {
        row.outcome = 'duplicate_key';
        row.detail = `Ambiguous key group: ${group.left.length} left row(s), ${group.right.length} right row(s). Every row is excluded from pairing; no first-wins or aggregation.`;
      }
      continue;
    }
    if (group.left.length !== 1 || group.right.length !== 1) continue;
    const left = group.left[0]!;
    const right = group.right[0]!;
    left.counterpart = {
      sheet: right.sheet,
      sourceRow: right.sourceRow,
      sourceRange: { ...right.sourceRange },
    };
    right.counterpart = {
      sheet: left.sheet,
      sourceRow: left.sourceRow,
      sourceRange: { ...left.sourceRange },
    };
    let outcome: 'matched' | 'invalid_amount' | 'amount_mismatch' = 'matched';
    let detail = 'Unique one-to-one key match (amounts not configured).';
    if (hasAmounts) {
      if (left.amount!.value === null || right.amount!.value === null) {
        outcome = 'invalid_amount';
        left.deltaReason = right.deltaReason = 'invalid_amount';
        detail = `Unique key, but amount excluded: left ${left.amount!.reason ?? 'numeric'}, right ${right.amount!.reason ?? 'numeric'}. Invalid amounts are not zero.`;
        pairedDeltas.add(null);
      } else {
        const delta = left.amount!.value - right.amount!.value;
        const finite = Number.isFinite(delta);
        left.delta = right.delta = finite ? delta : null;
        left.deltaReason = right.deltaReason = finite ? null : 'overflow';
        outcome = finite && Math.abs(delta) <= args.tolerance ? 'matched' : 'amount_mismatch';
        detail = finite
          ? `Unique key; signed left-minus-right delta ${delta} ${outcome === 'matched' ? 'is within' : 'exceeds'} absolute tolerance ${args.tolerance}.`
          : 'Unique key; signed amount difference exceeds finite numeric range. Classified as a mismatch, never a tolerance match.';
        pairedDeltas.add(finite ? delta : null);
        if (outcome === 'matched') matchedDeltas.add(delta);
      }
    }
    left.outcome = right.outcome = outcome;
    left.detail = right.detail = detail;
    pairs[outcome] += 1;
  }

  const sideCounts = (side: ReconciliationSide): ReconciliationSideCounts => {
    const outcomes = Object.fromEntries(
      RECONCILIATION_OUTCOMES.map((outcome) => [outcome, 0]),
    ) as Record<ReconciliationOutcome, number>;
    let sourceRows = 0;
    for (const row of rows)
      if (row.side === side) {
        sourceRows += 1;
        outcomes[row.outcome] += 1;
      }
    return {
      sourceRows,
      outcomes,
      accountedRows: Object.values(outcomes).reduce((sum, count) => sum + count, 0),
    };
  };
  const left = sideCounts('left');
  const right = sideCounts('right');
  const counts = {
    left,
    right,
    pairs,
    duplicateGroups,
    exceptionRows: rows.length - left.outcomes.matched - right.outcomes.matched,
    rowAccountingComplete:
      left.accountedRows === getSheet(workbook, args.leftSheet)!.rows.length - args.leftHeaderRow &&
      right.accountedRows ===
        getSheet(workbook, args.rightSheet)!.rows.length - args.rightHeaderRow,
  };
  const amounts = hasAmounts
    ? {
        left: amountControls(rows, 'left'),
        right: amountControls(rows, 'right'),
        pairedDeltas: pairedDeltas.result(),
        matchedDeltas: matchedDeltas.result(),
      }
    : null;
  const warnings: ValidationIssue[] = [];
  if (hasAmounts) warnings.push(issue('same-unit-assumption', currency.assumption));
  if (args.keyNormalization !== 'exact')
    warnings.push(
      issue(
        'key-normalization',
        `Opt-in ${args.keyNormalization} normalization applies only to string keys; it can merge otherwise different keys. Types remain distinct.`,
      ),
    );
  if (duplicateGroups)
    warnings.push(
      issue(
        'duplicate-keys',
        `${duplicateGroups} duplicate key group(s); ${left.outcomes.duplicate_key + right.outcomes.duplicate_key} source rows are explicitly unpaired.`,
      ),
    );
  const invalidKeys = left.outcomes.invalid_key + right.outcomes.invalid_key;
  if (invalidKeys)
    warnings.push(
      issue(
        'invalid-keys',
        `${invalidKeys} source row(s) have blank, invalid, error, unsupported, or volatile keys, including completely blank rows.`,
      ),
    );
  const discrepancies = pairs.amount_mismatch + left.outcomes.left_only + right.outcomes.right_only;
  if (discrepancies)
    warnings.push(
      issue(
        'reconciliation-discrepancies',
        `${pairs.amount_mismatch} mismatching pair(s), ${left.outcomes.left_only} left-only row(s), ${right.outcomes.right_only} right-only row(s). These are data discrepancies, not a financial conclusion.`,
      ),
    );
  if (amounts) {
    const excluded = amounts.left.total.excludedCount + amounts.right.total.excludedCount;
    const converted =
      amounts.left.total.numericTextConversions + amounts.right.total.numericTextConversions;
    if (excluded)
      warnings.push(
        issue(
          'excluded-amounts',
          `${excluded} nonnumeric, missing, error, unsupported, or volatile amount(s) are excluded from numeric totals, not treated as zero.`,
        ),
      );
    if (converted)
      warnings.push(
        issue(
          'numeric-text-conversions',
          `${converted} conservative numeric-text amount(s) converted using the engine numeric rule; no arbitrary currency parsing.`,
        ),
      );
    const totals = [
      amounts.left.total,
      amounts.left.matched,
      amounts.left.exceptions,
      amounts.right.total,
      amounts.right.matched,
      amounts.right.exceptions,
      amounts.pairedDeltas,
      amounts.matchedDeltas,
    ];
    if (
      totals.some((total) => total.overflow || total.roundoffBound === null) ||
      rows.some((row) => row.deltaReason === 'overflow')
    ) {
      warnings.push(
        issue(
          'numeric-overflow',
          'Some totals, roundoff bounds, or signed differences exceed finite numeric range. Unrepresentable results are explicitly unavailable; accounting counts remain complete.',
        ),
      );
    }
  }
  return { args, sources, rows, counts, amounts, currency, warnings };
}
