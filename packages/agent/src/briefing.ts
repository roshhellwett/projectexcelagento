import {
  createDeterministicWorkbookValueReader,
  indexToColumn,
  columnToIndex,
  maxColumnCount,
  summarizeNumericValues,
  toNumericOrNull,
  type Cell,
  type CellType,
  type FormulaValue,
  type NumericSummary,
  type Sheet,
  type Workbook,
} from '@excel-agent/engine';

export const BRIEFING_LIMITS = Object.freeze({
  categories: 8,
  charts: 4,
  histogramBins: 8,
  qualityExamples: 12,
  changedCellExamples: 20,
});

export interface BriefingSourceRange {
  sheet: string;
  startRow: number;
  endRow: number;
  startColumn: string;
  endColumn: string;
  a1: string;
  rowCount: number;
  cellCount: number;
}

export interface BriefingCategory {
  label: string;
  /** Retains value type so boolean true and text "true" never merge. */
  valueType: 'string' | 'boolean' | 'date';
  count: number;
}

export interface BriefingColumn {
  column: string;
  header: string;
  headerSource: BriefingSourceRange;
  sourceRange: BriefingSourceRange | null;
  numeric: NumericSummary;
  formulaCount: number;
  formulaErrorCount: number;
  unsupportedFormulaCount: number;
  volatileFormulaCount: number;
  numericTextCount: number;
  categories: {
    totalCount: number;
    distinctCount: number;
    top: BriefingCategory[];
    omittedCount: number;
    omittedCategoryCount: number;
  };
}

export interface BriefingChart {
  id: string;
  kind: 'histogram' | 'categorical';
  column: string;
  header: string;
  sourceRange: BriefingSourceRange;
  series: { label: string; count: number; lowerBound?: number; upperBound?: number }[];
  includedCount: number;
  omittedCount: number;
  notes: string[];
}

export interface AnalystBriefing {
  status: 'ready' | 'empty' | 'sheet-not-found';
  sheetName: string;
  headerRow: 1;
  dateSystem: '1900' | '1904';
  sourceRange: BriefingSourceRange | null;
  dataRange: BriefingSourceRange | null;
  dataRowCount: number;
  columnCount: number;
  columns: BriefingColumn[];
  charts: BriefingChart[];
  omittedChartCount: number;
  quality: {
    totalCells: number;
    missingCount: number;
    numericCount: number;
    nonNumericCount: number;
    formulaCount: number;
    formulaErrorCount: number;
    unsupportedFormulaCount: number;
    volatileFormulaCount: number;
    numericTextCount: number;
    raggedRowCount: number;
    fullyBlankRowCount: number;
    blankHeaderCount: number;
    duplicateHeaders: string[];
    completenessPercent: number | null;
    examples: { address: string; reason: string }[];
    omittedExampleCount: number;
  };
  findings: string[];
  limits: string[];
}

function range(sheet: Sheet, startRow: number, endRow: number, from: number, to: number) {
  if (endRow < startRow || to < from) return null;
  const startColumn = indexToColumn(from);
  const endColumn = indexToColumn(to);
  return {
    sheet: sheet.name,
    startRow,
    endRow,
    startColumn,
    endColumn,
    a1: `'${sheet.name.replace(/'/g, "''")}'!${startColumn}${startRow}:${endColumn}${endRow}`,
    rowCount: endRow - startRow + 1,
    cellCount: (endRow - startRow + 1) * (to - from + 1),
  } satisfies BriefingSourceRange;
}

const isMissing = (value: unknown) =>
  value === null || value === undefined || (typeof value === 'string' && value.trim() === '');
const lexicalOrder = (left: string, right: string) => (left < right ? -1 : left > right ? 1 : 0);
function category(value: FormulaValue): Omit<BriefingCategory, 'count'> | null {
  if (value instanceof Date) {
    return Number.isFinite(value.getTime())
      ? { label: value.toISOString(), valueType: 'date' }
      : null;
  }
  if (typeof value === 'string') return { label: value, valueType: 'string' };
  if (typeof value === 'boolean') return { label: String(value), valueType: 'boolean' };
  return null;
}

function histogram(column: BriefingColumn, numbers: number[]): BriefingChart | null {
  const { min, max } = column.numeric;
  if (min === null || max === null || !column.sourceRange) return null;
  const span = max - min;
  if (!Number.isFinite(span)) return null;
  const binCount = min === max ? 1 : BRIEFING_LIMITS.histogramBins;
  const series = Array.from({ length: binCount }, (_, index) => {
    const lowerBound = min + span * (index / binCount);
    const upperBound = index === binCount - 1 ? max : min + span * ((index + 1) / binCount);
    return {
      label:
        min === max
          ? String(min)
          : `[${lowerBound}, ${upperBound}${index === binCount - 1 ? ']' : ')'}`,
      count: 0,
      lowerBound,
      upperBound,
    };
  });
  for (const number of numbers) {
    const index =
      span === 0 ? 0 : Math.min(binCount - 1, Math.floor(((number - min) / span) * binCount));
    series[index]!.count += 1;
  }
  return {
    id: `numeric-${column.column}`,
    kind: 'histogram',
    column: column.column,
    header: column.header,
    sourceRange: column.sourceRange,
    series,
    includedCount: numbers.length,
    omittedCount: 0,
    notes: [
      'Equal-width numeric bins; lower bounds included, upper bounds excluded except the last bin. Missing and nonnumeric cells excluded.',
    ],
  };
}

/** Full-sheet, read-only descriptive briefing. Row 1 is explicitly assumed to be the header. */
export function createAnalystBriefing(workbook: Workbook, sheetName: string): AnalystBriefing {
  const sheet = workbook.sheets.find(
    (item) => item.name.toLowerCase() === sheetName.trim().toLowerCase(),
  );
  const width = sheet ? maxColumnCount(sheet.rows) : 0;
  const rowCount = Math.max(0, (sheet?.rows.length ?? 0) - 1);
  const briefing: AnalystBriefing = {
    status: !sheet ? 'sheet-not-found' : rowCount === 0 || width === 0 ? 'empty' : 'ready',
    sheetName: sheet?.name ?? sheetName,
    headerRow: 1,
    dateSystem: workbook.dateSystem ?? '1900',
    sourceRange: sheet ? range(sheet, 1, sheet.rows.length, 0, width - 1) : null,
    dataRange: sheet ? range(sheet, 2, sheet.rows.length, 0, width - 1) : null,
    dataRowCount: rowCount,
    columnCount: width,
    columns: [],
    charts: [],
    omittedChartCount: 0,
    quality: {
      totalCells: rowCount * width,
      missingCount: 0,
      numericCount: 0,
      nonNumericCount: 0,
      formulaCount: 0,
      formulaErrorCount: 0,
      unsupportedFormulaCount: 0,
      volatileFormulaCount: 0,
      numericTextCount: 0,
      raggedRowCount: 0,
      fullyBlankRowCount: 0,
      blankHeaderCount: 0,
      duplicateHeaders: [],
      completenessPercent: null,
      examples: [],
      omittedExampleCount: 0,
    },
    findings: [],
    limits: [
      'Row 1 is assumed to be the header; every subsequent stored row is included, including blank and ragged rows. Source ranges are not samples.',
      'Numeric text is included using engine statistics rules. Booleans and dates are nonnumeric; errors and unsupported or clock/random formulas (including their dependents) are excluded, never replaced by formula caches.',
      'Counts and statistics scan full source ranges. Charts are bounded previews, not time series or forecasts; column units are not inferred and totals across columns are not combined.',
      `At most ${BRIEFING_LIMITS.charts} charts, ${BRIEFING_LIMITS.categories} categories per chart, and ${BRIEFING_LIMITS.qualityExamples} quality examples are shown.`,
      'Descriptive data only: no causal, financial-performance, or forecasting claims. Completeness measures nonblank cells, not correctness.',
    ],
  };
  if (!sheet) {
    briefing.findings.push(`Sheet "${sheetName}" was not found.`);
    return briefing;
  }
  if (sheet.rows.length === 0 || width === 0) {
    briefing.findings.push('No stored cells are available for this sheet.');
    return briefing;
  }
  const read = createDeterministicWorkbookValueReader(workbook);
  const candidates: BriefingChart[] = [];
  const headerCounts = new Map<string, number>();
  const blankRows = new Array<boolean>(rowCount).fill(true);
  let issueCount = 0;
  const issue = (column: string, row: number, reason: string) => {
    issueCount += 1;
    if (briefing.quality.examples.length < BRIEFING_LIMITS.qualityExamples)
      briefing.quality.examples.push({
        address: range(sheet, row, row, columnToIndex(column)!, columnToIndex(column)!)!.a1,
        reason,
      });
  };
  for (let index = 0; index < width; index += 1) {
    const letter = indexToColumn(index);
    const headerValue = read(sheet.name, index, 1);
    const header =
      headerValue.reason || isMissing(headerValue.value)
        ? `Column ${letter}`
        : headerValue.value instanceof Date
          ? (category(headerValue.value)?.label ?? `Column ${letter}`)
          : String(headerValue.value);
    if (isMissing(headerValue.value)) briefing.quality.blankHeaderCount += 1;
    else if (!headerValue.reason) headerCounts.set(header, (headerCounts.get(header) ?? 0) + 1);
    const values: FormulaValue[] = [];
    const numbers: number[] = [];
    const categories = new Map<string, BriefingCategory>();
    const column: BriefingColumn = {
      column: letter,
      header,
      headerSource: range(sheet, 1, 1, index, index)!,
      sourceRange: range(sheet, 2, sheet.rows.length, index, index),
      numeric: summarizeNumericValues([]),
      formulaCount: 0,
      formulaErrorCount: 0,
      unsupportedFormulaCount: 0,
      volatileFormulaCount: 0,
      numericTextCount: 0,
      categories: {
        totalCount: 0,
        distinctCount: 0,
        top: [],
        omittedCount: 0,
        omittedCategoryCount: 0,
      },
    };
    for (let row = 2; row <= sheet.rows.length; row += 1) {
      const { value, reason } = read(sheet.name, index, row);
      const cell = sheet.rows[row - 1]?.[index];
      values.push(value);
      if (cell?.formula || cell?.type === 'formula') column.formulaCount += 1;
      if (reason === 'unsupported') column.unsupportedFormulaCount += 1;
      else if (reason === 'volatile') column.volatileFormulaCount += 1;
      else if (reason === 'error') column.formulaErrorCount += 1;
      if (!isMissing(value)) blankRows[row - 2] = false;
      if (reason) {
        issue(
          letter,
          row,
          reason === 'error'
            ? `Error value ${value}; excluded`
            : `${reason === 'volatile' ? 'Clock/random' : 'Unsupported'} formula or dependency; excluded`,
        );
        continue;
      }
      if (isMissing(value)) {
        issue(letter, row, 'Missing value');
        continue;
      }
      const numeric = toNumericOrNull(value);
      if (numeric !== null) {
        numbers.push(numeric);
        if (typeof value === 'string') column.numericTextCount += 1;
        continue;
      }
      const entry = category(value);
      if (entry) {
        const id = JSON.stringify([entry.valueType, entry.label]);
        const existing = categories.get(id);
        categories.set(id, { ...entry, count: (existing?.count ?? 0) + 1 });
      } else issue(letter, row, 'Non-finite or invalid value; excluded');
    }
    column.numeric = summarizeNumericValues(values);
    const ranked = [...categories.values()].sort(
      (left, right) =>
        right.count - left.count ||
        lexicalOrder(left.label, right.label) ||
        lexicalOrder(left.valueType, right.valueType),
    );
    column.categories.top = ranked.slice(0, BRIEFING_LIMITS.categories);
    column.categories.totalCount = ranked.reduce((total, item) => total + item.count, 0);
    column.categories.distinctCount = ranked.length;
    column.categories.omittedCount =
      column.categories.totalCount -
      column.categories.top.reduce((total, item) => total + item.count, 0);
    column.categories.omittedCategoryCount = ranked.length - column.categories.top.length;
    briefing.columns.push(column);
    for (const name of ['missingCount', 'numericCount', 'nonNumericCount'] as const)
      briefing.quality[name] += column.numeric[name];
    for (const name of [
      'formulaCount',
      'formulaErrorCount',
      'unsupportedFormulaCount',
      'volatileFormulaCount',
      'numericTextCount',
    ] as const)
      briefing.quality[name] += column[name];
    const numericChart = histogram(column, numbers);
    if (numericChart) candidates.push(numericChart);
    else if (numbers.length && column.sourceRange)
      briefing.findings.push(
        `${column.header} (${letter}): histogram omitted because the numeric range width exceeds the finite numeric range. Statistics and exact exclusions remain in the column summary.`,
      );
    if (column.categories.top.length && column.sourceRange)
      candidates.push({
        id: `categories-${letter}`,
        kind: 'categorical',
        column: letter,
        header,
        sourceRange: column.sourceRange,
        series: column.categories.top.map((item) => ({
          label: `${item.label} (${item.valueType})`,
          count: item.count,
        })),
        includedCount: column.categories.totalCount - column.categories.omittedCount,
        omittedCount: column.categories.omittedCount,
        notes: [
          `Top ${column.categories.top.length} of ${column.categories.distinctCount} exact, case-sensitive nonnumeric categories; ${column.categories.omittedCategoryCount} categories omitted. Numeric, missing, and error/unsupported cells excluded.`,
        ],
      });
  }
  briefing.charts = candidates.slice(0, BRIEFING_LIMITS.charts);
  briefing.omittedChartCount = candidates.length - briefing.charts.length;
  briefing.quality.raggedRowCount = sheet.rows.slice(1).filter((row) => row.length < width).length;
  briefing.quality.fullyBlankRowCount = blankRows.filter(Boolean).length;
  briefing.quality.duplicateHeaders = [...headerCounts]
    .filter(([, count]) => count > 1)
    .map(([header]) => header);
  briefing.quality.omittedExampleCount = issueCount - briefing.quality.examples.length;
  briefing.quality.completenessPercent =
    briefing.quality.totalCells > 0
      ? (1 - briefing.quality.missingCount / briefing.quality.totalCells) * 100
      : null;
  const quality = briefing.quality;
  briefing.findings.push(
    `${rowCount} data rows and ${width} columns inspected after the assumed header.`,
  );
  if (quality.missingCount)
    briefing.findings.push(
      `${quality.missingCount} missing cells; ${quality.fullyBlankRowCount} fully blank data rows.`,
    );
  if (quality.nonNumericCount)
    briefing.findings.push(
      `${quality.nonNumericCount} nonnumeric cells excluded from numeric statistics (includes errors and unavailable formulas).`,
    );
  if (quality.formulaErrorCount + quality.unsupportedFormulaCount + quality.volatileFormulaCount)
    briefing.findings.push(
      `${quality.formulaErrorCount} error values, ${quality.unsupportedFormulaCount} unsupported formulas/dependents, and ${quality.volatileFormulaCount} clock/random formulas/dependents excluded.`,
    );
  if (quality.blankHeaderCount || quality.duplicateHeaders.length)
    briefing.findings.push(
      `${quality.blankHeaderCount} blank headers and ${quality.duplicateHeaders.length} repeated header labels; columns are identified by letters, not inferred meaning.`,
    );
  if (!quality.numericCount)
    briefing.findings.push(
      'No numeric observations; numeric totals and distributions are undefined, not zero.',
    );
  if (briefing.status === 'empty')
    briefing.findings.push('No data observations below the assumed header.');
  return briefing;
}

export type BaselineValue =
  | string
  | number
  | boolean
  | null
  | {
      kind: 'date';
      epochMilliseconds: number | null;
      iso: string | null;
    }
  | { kind: 'non-finite-number'; value: string };

export interface BaselineCellSnapshot {
  present: boolean;
  type?: CellType;
  value?: BaselineValue;
  formula?: string | null;
  numberFormat?: string | null;
}

export interface WorkbookBaselineComparison {
  baselineDateSystem: '1900' | '1904';
  currentDateSystem: '1900' | '1904';
  dateSystemChanged: boolean;
  sheets: {
    sheetName: string;
    presence: 'both' | 'added' | 'removed';
    baselineRows: number;
    currentRows: number;
    rowDelta: number;
    baselineColumns: number;
    currentColumns: number;
    columnDelta: number;
    baselineSourceRange: BriefingSourceRange | null;
    currentSourceRange: BriefingSourceRange | null;
    headerChanged: boolean;
    layoutChanged: boolean;
    changedCellCount: number;
  }[];
  changedCellCount: number;
  changedCells: {
    sheet: string;
    address: string;
    before: BaselineCellSnapshot;
    after: BaselineCellSnapshot;
  }[];
  omittedChangedCellCount: number;
  notes: string[];
}

function snapshot(cell: Cell | undefined): BaselineCellSnapshot {
  if (!cell) return { present: false };
  let value: BaselineValue = cell.value instanceof Date ? null : cell.value;
  if (cell.value instanceof Date) {
    const milliseconds = cell.value.getTime();
    value = {
      kind: 'date',
      epochMilliseconds: Number.isFinite(milliseconds) ? milliseconds : null,
      iso: Number.isFinite(milliseconds) ? cell.value.toISOString() : null,
    };
  } else if (typeof cell.value === 'number' && !Number.isFinite(cell.value)) {
    value = { kind: 'non-finite-number', value: String(cell.value) };
  }
  return {
    present: true,
    type: cell.type,
    value,
    formula: cell.formula ?? null,
    numberFormat: cell.numberFormat ?? null,
  };
}

const sameSnapshot = (left: BaselineCellSnapshot, right: BaselineCellSnapshot) => {
  // Compare numbers separately: Object.is retains exact -0 versus +0, unlike JSON.
  if (
    typeof left.value === 'number' &&
    typeof right.value === 'number' &&
    !Object.is(left.value, right.value)
  )
    return false;
  return JSON.stringify(left) === JSON.stringify(right);
};

/** Exact stored-cell and shape comparison, not row matching or comparable aggregate inference.
 * Formula source/cache/type/format and UTC Date instants are compared. Epoch changes are explicit
 * workbook metadata changes, not falsely counted as edits to every otherwise identical cell. */
export function compareWorkbookBaseline(
  baseline: Workbook,
  current: Workbook,
): WorkbookBaselineComparison {
  const result: WorkbookBaselineComparison = {
    baselineDateSystem: baseline.dateSystem ?? '1900',
    currentDateSystem: current.dateSystem ?? '1900',
    dateSystemChanged: (baseline.dateSystem ?? '1900') !== (current.dateSystem ?? '1900'),
    sheets: [],
    changedCellCount: 0,
    changedCells: [],
    omittedChangedCellCount: 0,
    notes: [
      'Address-based comparison of exact stored cells, including headers, formulas, caches, types, number formats, and Date instants. Absent cells differ from explicit blank cells.',
      'No record alignment, trends, or comparable aggregate deltas are inferred, even when headers match. Shifted headers or rows may create many positional changes.',
      `Changed-cell examples are bounded to ${BRIEFING_LIMITS.changedCellExamples}; change counts scan all stored addresses.`,
    ],
  };
  const names = [
    ...new Set([
      ...baseline.sheets.map((sheet) => sheet.name),
      ...current.sheets.map((sheet) => sheet.name),
    ]),
  ];
  for (const name of names) {
    const before = baseline.sheets.find((sheet) => sheet.name === name);
    const after = current.sheets.find((sheet) => sheet.name === name);
    const beforeWidth = before ? maxColumnCount(before.rows) : 0;
    const afterWidth = after ? maxColumnCount(after.rows) : 0;
    const beforeRows = before?.rows.length ?? 0;
    const afterRows = after?.rows.length ?? 0;
    const headerChanged = Array.from(
      { length: Math.max(beforeWidth, afterWidth) },
      (_, index) => index,
    ).some(
      (index) =>
        !sameSnapshot(snapshot(before?.rows[0]?.[index]), snapshot(after?.rows[0]?.[index])),
    );
    const sheet = {
      sheetName: name,
      presence: !before ? ('added' as const) : !after ? ('removed' as const) : ('both' as const),
      baselineRows: beforeRows,
      currentRows: afterRows,
      rowDelta: afterRows - beforeRows,
      baselineColumns: beforeWidth,
      currentColumns: afterWidth,
      columnDelta: afterWidth - beforeWidth,
      baselineSourceRange: before ? range(before, 1, beforeRows, 0, beforeWidth - 1) : null,
      currentSourceRange: after ? range(after, 1, afterRows, 0, afterWidth - 1) : null,
      headerChanged,
      layoutChanged:
        !before ||
        !after ||
        headerChanged ||
        beforeRows !== afterRows ||
        beforeWidth !== afterWidth ||
        before.rows.some((row, index) => row.length !== after.rows[index]?.length),
      changedCellCount: 0,
    };
    for (let row = 0; row < Math.max(beforeRows, afterRows); row += 1) {
      const rowWidth = Math.max(before?.rows[row]?.length ?? 0, after?.rows[row]?.length ?? 0);
      for (let column = 0; column < rowWidth; column += 1) {
        const oldCell = snapshot(before?.rows[row]?.[column]);
        const newCell = snapshot(after?.rows[row]?.[column]);
        if (sameSnapshot(oldCell, newCell)) continue;
        sheet.changedCellCount += 1;
        result.changedCellCount += 1;
        if (result.changedCells.length < BRIEFING_LIMITS.changedCellExamples)
          result.changedCells.push({
            sheet: name,
            address: range(before ?? after!, row + 1, row + 1, column, column)!.a1,
            before: oldCell,
            after: newCell,
          });
      }
    }
    result.sheets.push(sheet);
  }
  result.omittedChangedCellCount = result.changedCellCount - result.changedCells.length;
  if (result.dateSystemChanged)
    result.notes.push(
      'The workbook date epoch changed. Identical serials and formulas may now mean different dates; evaluated formula changes are not inferred from stored-cell comparison.',
    );
  if (result.sheets.some((sheet) => sheet.layoutChanged))
    result.notes.push(
      'Layouts differ: treat changed-cell examples as positional evidence only, not like-for-like metric changes.',
    );
  return result;
}
