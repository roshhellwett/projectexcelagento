import {
  columnToIndex,
  indexToColumn,
  type Cell,
  type Sheet,
  type Workbook,
} from '@excel-agent/engine';
import { getColumnProfiles } from './analysis.js';
import type { ToolDefinition } from './types.js';

function sheetMaxCols(rows: Cell[][]): number {
  return rows.reduce((max, r) => Math.max(max, r.length), 0);
}

export interface WorkbookOverviewResult {
  sheets: {
    name: string;
    rowCount: number;
    columnCount: number;
    headers: string[];
  }[];
}

export interface ColumnProfileResult {
  sheet: string;
  column: string;
  headerName: string;
  inferredType: 'numeric' | 'date' | 'text' | 'empty';
  totalRows: number;
  nonBlankCount: number;
  blankCount: number;
  distinctCount: number;
  sampleValues: unknown[];
  sum?: number;
  average?: number;
  min?: number;
  max?: number;
}

export interface CellRangeResult {
  sheet: string;
  startRow: number;
  endRow: number;
  startColumn: string;
  endColumn: string;
  rows: { rowNumber: number; cells: Record<string, unknown> }[];
}

export interface SearchSheetResult {
  sheet: string;
  query: string;
  matches: {
    cell: string;
    rowNumber: number;
    columnLetter: string;
    value: unknown;
    rowContext: Record<string, unknown>;
  }[];
  totalMatches: number;
}

export interface AggregateResult {
  sheet: string;
  column: string;
  metric: 'sum' | 'avg' | 'min' | 'max' | 'count' | 'count_distinct';
  value: number | null;
  count: number;
}

/** Get high-level metadata across all sheets in the workbook. */
export function getWorkbookOverview(workbook: Workbook): WorkbookOverviewResult {
  return {
    sheets: workbook.sheets.map((sheet) => {
      const headerRow = sheet.rows[0] ?? [];
      const headers = headerRow.map((cell, idx) => {
        const val = cell?.value;
        return val !== null && val !== undefined && String(val).trim() !== ''
          ? String(val)
          : `Col ${indexToColumn(idx)}`;
      });
      return {
        name: sheet.name,
        rowCount: sheet.rows.length,
        columnCount: sheetMaxCols(sheet.rows),
        headers,
      };
    }),
  };
}

function resolveSheet(workbook: Workbook, sheetName?: string): Sheet | undefined {
  if (!sheetName) return workbook.sheets[0];
  return workbook.sheets.find((s) => s.name.toLowerCase() === sheetName.trim().toLowerCase());
}

/** Profile a single column for deep analysis, cardinality, distributions, and aggregates. */
export function profileColumn(
  workbook: Workbook,
  sheetName: string | undefined,
  column: string,
): ColumnProfileResult | { error: string } {
  const sheet = resolveSheet(workbook, sheetName);
  if (!sheet) return { error: `Sheet "${sheetName ?? ''}" not found.` };

  const targetIdx = columnToIndex(column);
  if (targetIdx === undefined) return { error: `Invalid column reference: "${column}".` };

  const profiles = getColumnProfiles(sheet);
  const profile = profiles.find((p) => p.letter.toUpperCase() === column.toUpperCase());
  if (!profile) return { error: `Column ${column} not found in sheet "${sheet.name}".` };

  const totalRows = Math.max(0, sheet.rows.length - 1);
  const blankCount = Math.max(0, totalRows - profile.nonBlankCount);

  return {
    sheet: sheet.name,
    column: profile.letter,
    headerName: profile.rawName,
    inferredType: profile.isNumeric
      ? 'numeric'
      : profile.isDate
        ? 'date'
        : profile.nonBlankCount === 0
          ? 'empty'
          : 'text',
    totalRows,
    nonBlankCount: profile.nonBlankCount,
    blankCount,
    distinctCount: profile.distinct.size,
    sampleValues: Array.from(profile.distinct.keys()).slice(0, 8),
    ...(profile.isNumeric && profile.sum !== undefined
      ? {
          sum: Math.round(profile.sum * 1000) / 1000,
          average: Math.round((profile.avg ?? 0) * 1000) / 1000,
          min: profile.min,
          max: profile.max,
        }
      : {}),
  };
}

/** Read a bounded cell range with headers and values. */
export function readCellRange(
  workbook: Workbook,
  sheetName: string | undefined,
  startRow = 1,
  endRow = 15,
  startColumn = 'A',
  endColumn?: string,
): CellRangeResult | { error: string } {
  const sheet = resolveSheet(workbook, sheetName);
  if (!sheet) return { error: `Sheet "${sheetName ?? ''}" not found.` };

  const startColIdx = columnToIndex(startColumn) ?? 0;
  const maxCol = sheetMaxCols(sheet.rows);
  const endColIdx = endColumn
    ? (columnToIndex(endColumn) ?? maxCol - 1)
    : Math.min(startColIdx + 10, maxCol - 1);

  const boundedStartRow = Math.max(1, startRow);
  const boundedEndRow = Math.min(
    Math.max(boundedStartRow, endRow),
    sheet.rows.length,
    boundedStartRow + 50,
  );

  const headers = (sheet.rows[0] ?? []).map((c, i) => String(c?.value ?? indexToColumn(i)));

  const rows: CellRangeResult['rows'] = [];
  for (let r = boundedStartRow; r <= boundedEndRow; r += 1) {
    const rowCells = sheet.rows[r - 1] ?? [];
    const record: Record<string, unknown> = {};
    for (let c = startColIdx; c <= endColIdx; c += 1) {
      const colLetter = indexToColumn(c);
      const colHeader = headers[c] || colLetter;
      const cell = rowCells[c];
      record[`${colLetter} (${colHeader})`] = cell?.formula
        ? `=${cell.formula}`
        : (cell?.value ?? null);
    }
    rows.push({ rowNumber: r, cells: record });
  }

  return {
    sheet: sheet.name,
    startRow: boundedStartRow,
    endRow: boundedEndRow,
    startColumn: indexToColumn(startColIdx),
    endColumn: indexToColumn(endColIdx),
    rows,
  };
}

/** Search sheet for query text or numbers. */
export function searchSheet(
  workbook: Workbook,
  sheetName: string | undefined,
  query: string,
  limit = 20,
): SearchSheetResult | { error: string } {
  const sheet = resolveSheet(workbook, sheetName);
  if (!sheet) return { error: `Sheet "${sheetName ?? ''}" not found.` };

  const needle = query.trim().toLowerCase();
  if (!needle) return { sheet: sheet.name, query, matches: [], totalMatches: 0 };

  const matches: SearchSheetResult['matches'] = [];
  const headers = (sheet.rows[0] ?? []).map((c, i) => String(c?.value ?? indexToColumn(i)));

  for (let r = 0; r < sheet.rows.length; r += 1) {
    const row = sheet.rows[r] ?? [];
    for (let c = 0; c < row.length; c += 1) {
      const cell = row[c];
      if (!cell || cell.value === null || cell.value === undefined) continue;
      const strVal = String(cell.value).toLowerCase();
      if (strVal.includes(needle)) {
        const colLetter = indexToColumn(c);
        const rowContext: Record<string, unknown> = {};
        row.forEach((cellItem, idx) => {
          rowContext[headers[idx] || indexToColumn(idx)] = cellItem?.value ?? null;
        });
        matches.push({
          cell: `${colLetter}${r + 1}`,
          rowNumber: r + 1,
          columnLetter: colLetter,
          value: cell.value,
          rowContext,
        });
        if (matches.length >= limit) break;
      }
    }
    if (matches.length >= limit) break;
  }

  return {
    sheet: sheet.name,
    query,
    matches,
    totalMatches: matches.length,
  };
}

/** Compute deterministic aggregates on any numeric column. */
export function calculateAggregate(
  workbook: Workbook,
  sheetName: string | undefined,
  column: string,
  metric: 'sum' | 'avg' | 'min' | 'max' | 'count' | 'count_distinct',
): AggregateResult | { error: string } {
  const profileRes = profileColumn(workbook, sheetName, column);
  if ('error' in profileRes) return profileRes;

  const sheet = resolveSheet(workbook, sheetName)!;
  const colIdx = columnToIndex(column)!;
  const values: number[] = [];
  const allValues = new Set<unknown>();

  for (let r = 1; r < sheet.rows.length; r += 1) {
    const cell = sheet.rows[r]?.[colIdx];
    if (cell && cell.value !== null && cell.value !== undefined && cell.value !== '') {
      allValues.add(cell.value);
      const num = typeof cell.value === 'number' ? cell.value : Number(cell.value);
      if (!isNaN(num)) values.push(num);
    }
  }

  let resultValue: number | null = null;
  if (metric === 'count') {
    resultValue = values.length;
  } else if (metric === 'count_distinct') {
    resultValue = allValues.size;
  } else if (values.length > 0) {
    switch (metric) {
      case 'sum':
        resultValue = values.reduce((a, b) => a + b, 0);
        break;
      case 'avg':
        resultValue = values.reduce((a, b) => a + b, 0) / values.length;
        break;
      case 'min':
        resultValue = values.reduce((lowest, value) => Math.min(lowest, value), Infinity);
        break;
      case 'max':
        resultValue = values.reduce((highest, value) => Math.max(highest, value), -Infinity);
        break;
    }
  }

  return {
    sheet: sheet.name,
    column: column.toUpperCase(),
    metric,
    value: resultValue !== null ? Math.round(resultValue * 1000) / 1000 : null,
    count: values.length,
  };
}

/** Tool definitions for LLM function calling */
export const READ_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    type: 'function',
    function: {
      name: 'get_workbook_overview',
      description:
        'Get an overview of all sheets, row counts, and column headers in the current workbook.',
      parameters: {
        type: 'object',
        properties: {},
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'profile_column',
      description:
        'Get deep statistical profiling on a column: data type, nulls, distinct values, and numerical aggregates (sum, avg, min, max).',
      parameters: {
        type: 'object',
        properties: {
          sheet: { type: 'string', description: 'Sheet name (defaults to active sheet)' },
          column: { type: 'string', description: 'Column letter, e.g. "A", "C", "D"' },
        },
        required: ['column'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'read_cell_range',
      description: 'Read the actual data and formulas from a bounded range of rows and columns.',
      parameters: {
        type: 'object',
        properties: {
          sheet: { type: 'string', description: 'Sheet name' },
          startRow: { type: 'number', description: 'Starting row number (1-based, default 1)' },
          endRow: { type: 'number', description: 'Ending row number (default 15)' },
          startColumn: { type: 'string', description: 'Start column letter (default "A")' },
          endColumn: { type: 'string', description: 'End column letter' },
        },
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'search_sheet',
      description: 'Search for text, customer names, IDs, or numbers across cells in the sheet.',
      parameters: {
        type: 'object',
        properties: {
          sheet: { type: 'string', description: 'Sheet name' },
          query: { type: 'string', description: 'Text or number to search for' },
          limit: { type: 'number', description: 'Maximum matches to return (default 15)' },
        },
        required: ['query'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'calculate_aggregate',
      description:
        'Calculate an exact, deterministic aggregate (sum, avg, min, max, count, count_distinct) on a column.',
      parameters: {
        type: 'object',
        properties: {
          sheet: { type: 'string', description: 'Sheet name' },
          column: { type: 'string', description: 'Column letter, e.g. "E"' },
          metric: {
            type: 'string',
            enum: ['sum', 'avg', 'min', 'max', 'count', 'count_distinct'],
            description: 'The aggregate function to compute',
          },
        },
        required: ['column', 'metric'],
      },
    },
  },
];
