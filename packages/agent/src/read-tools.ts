import {
  columnToIndex,
  aggregateCells,
  createCell,
  toNumericOrNull,
  createWorkbookValueReader,
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

function safeInteger(value: number, fallback: number): number {
  return Number.isFinite(value) ? Math.floor(value) : fallback;
}

function safeLimit(value: number, fallback: number, maximum: number): number {
  return Math.min(maximum, Math.max(0, safeInteger(value, fallback)));
}

const MAX_READ_COLUMNS = 100;

/** Keep duplicate or special header names from overwriting cells in a tool result. */
function recordHeaders(headerRow: Cell[]): string[] {
  const labels = headerRow.map(
    (cell, index) => String(cell?.value ?? '').trim() || indexToColumn(index),
  );
  const counts = new Map<string, number>();
  for (const label of labels) counts.set(label, (counts.get(label) ?? 0) + 1);
  const used = new Set<string>();
  return labels.map((label, index) => {
    let key = counts.get(label) === 1 ? label : `${label} (${indexToColumn(index)})`;
    while (used.has(key)) key = `${key} (${indexToColumn(index)})`;
    used.add(key);
    return key;
  });
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

  const read = createWorkbookValueReader(workbook);
  const profiles = getColumnProfiles(sheet, (column, row) => read(sheet.name, column, row));
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
  endRow = 50,
  startColumn = 'A',
  endColumn?: string,
): CellRangeResult | { error: string } {
  const sheet = resolveSheet(workbook, sheetName);
  if (!sheet) return { error: `Sheet "${sheetName ?? ''}" not found.` };

  const startColIdx = columnToIndex(startColumn);
  if (startColIdx === undefined) {
    return { error: `Invalid start column reference: "${startColumn}".` };
  }
  const requestedEndColIdx = endColumn === undefined ? undefined : columnToIndex(endColumn);
  if (endColumn !== undefined && requestedEndColIdx === undefined) {
    return { error: `Invalid end column reference: "${endColumn}".` };
  }
  if (requestedEndColIdx !== undefined && requestedEndColIdx < startColIdx) {
    return { error: 'The end column must not come before the start column.' };
  }
  const maxCol = sheetMaxCols(sheet.rows);
  if (maxCol === 0) {
    const safeStartRow = Math.min(
      Math.max(1, safeInteger(startRow, 1)),
      Math.max(1, sheet.rows.length),
    );
    return {
      sheet: sheet.name,
      startRow: safeStartRow,
      endRow: safeStartRow,
      startColumn: 'A',
      endColumn: 'A',
      rows: [],
    };
  }
  if (startColIdx >= maxCol) {
    return { error: `Column "${startColumn}" not found in sheet "${sheet.name}".` };
  }
  const endColIdx = Math.min(
    requestedEndColIdx ?? startColIdx + 49,
    startColIdx + MAX_READ_COLUMNS - 1,
    maxCol - 1,
  );

  const boundedStartRow = Math.min(
    Math.max(1, safeInteger(startRow, 1)),
    Math.max(1, sheet.rows.length),
  );
  const boundedEndRow = Math.min(
    Math.max(boundedStartRow, safeInteger(endRow, boundedStartRow + 49)),
    sheet.rows.length,
    boundedStartRow + 999,
  );

  // Intelligently identify the best header row across the top 10 rows (handles sheets with title rows)
  let headerRowIndex = 0;
  let maxHeaderCount = 0;
  for (let r = 0; r < Math.min(10, sheet.rows.length); r += 1) {
    const rCells = sheet.rows[r] ?? [];
    let count = 0;
    for (const cell of rCells) {
      const val = String(cell?.value ?? '').trim();
      if (val && (/^fy\s*'?\d{2,4}$/i.test(val) || /^20\d{2}$/.test(val) || val.length > 1)) {
        count += 1;
      }
    }
    if (count > maxHeaderCount) {
      maxHeaderCount = count;
      headerRowIndex = r;
    }
  }

  const headers = (sheet.rows[headerRowIndex] ?? []).map(
    (c, i) => String(c?.value ?? '').trim() || indexToColumn(i),
  );

  const rows: CellRangeResult['rows'] = [];
  for (let r = boundedStartRow; r <= boundedEndRow; r += 1) {
    const rowCells = sheet.rows[r - 1] ?? [];
    const record: Record<string, unknown> = {};
    for (let c = startColIdx; c <= endColIdx; c += 1) {
      const colLetter = indexToColumn(c);
      const colHeader = headers[c] || colLetter;
      const cell = rowCells[c];
      record[`${colLetter} (${colHeader})`] = cell?.formula
        ? `=${cell.formula.replace(/^=/, '')}`
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
  limit = 100,
): SearchSheetResult | { error: string } {
  const sheet = resolveSheet(workbook, sheetName);
  if (!sheet) return { error: `Sheet "${sheetName ?? ''}" not found.` };

  const needle = query.trim().toLowerCase();
  const boundedLimit = safeLimit(limit, 100, 500);
  if (!needle) return { sheet: sheet.name, query, matches: [], totalMatches: 0 };

  const matches: SearchSheetResult['matches'] = [];

  let headerRowIndex = 0;
  let maxHeaderCount = 0;
  for (let r = 0; r < Math.min(10, sheet.rows.length); r += 1) {
    const rCells = sheet.rows[r] ?? [];
    let count = 0;
    for (const cell of rCells) {
      const val = String(cell?.value ?? '').trim();
      if (val && (/^fy\s*'?\d{2,4}$/i.test(val) || /^20\d{2}$/.test(val) || val.length > 1)) {
        count += 1;
      }
    }
    if (count > maxHeaderCount) {
      maxHeaderCount = count;
      headerRowIndex = r;
    }
  }

  const headers = recordHeaders(sheet.rows[headerRowIndex] ?? []);
  let totalMatches = 0;
  const read = createWorkbookValueReader(workbook);

  for (let r = 0; r < sheet.rows.length; r += 1) {
    const row = sheet.rows[r] ?? [];
    for (let c = 0; c < row.length; c += 1) {
      const cell = row[c];
      if (!cell) continue;
      const value = read(sheet.name, c, r + 1);
      if (value === null) continue;
      const strVal = String(value).toLowerCase();
      if (strVal.includes(needle)) {
        totalMatches += 1;
        if (matches.length < boundedLimit) {
          const colLetter = indexToColumn(c);
          const rowContext: Record<string, unknown> = Object.fromEntries(
            row.map((_cell, idx) => [
              headers[idx] || indexToColumn(idx),
              read(sheet.name, idx, r + 1),
            ]),
          );
          matches.push({
            cell: `${colLetter}${r + 1}`,
            rowNumber: r + 1,
            columnLetter: colLetter,
            value,
            rowContext,
          });
        }
      }
    }
  }

  return {
    sheet: sheet.name,
    query,
    matches,
    totalMatches,
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
  const cells: Cell[] = [];
  const read = createWorkbookValueReader(workbook);

  for (let r = 1; r < sheet.rows.length; r += 1) {
    cells.push(createCell(read(sheet.name, colIdx, r + 1)));
  }

  if (!['sum', 'avg', 'min', 'max', 'count', 'count_distinct'].includes(metric)) {
    return { error: `Unknown aggregate metric "${metric}".` };
  }
  const outcome = aggregateCells(cells, metric === 'avg' ? 'average' : metric);
  const numericCount = cells.reduce(
    (count, cell) => count + (toNumericOrNull(cell.value) === null ? 0 : 1),
    0,
  );
  const resultValue =
    metric === 'count' || metric === 'count_distinct' || numericCount > 0 || metric === 'sum'
      ? outcome.value
      : null;

  return {
    sheet: sheet.name,
    column: column.toUpperCase(),
    metric,
    value: resultValue,
    count: numericCount,
  };
}

export interface QuerySheetCondition {
  column?: string;
  header?: string;
  operator?: 'equals' | 'contains' | 'startsWith' | 'endsWith' | 'gt' | 'lt';
  value: string | number;
}

export interface QuerySheetRecordsResult {
  sheet: string;
  totalMatchingRows: number;
  totalSheetRows: number;
  matchingRowNumbers: number[];
  sampleMatchingRows: { rowNumber: number; cells: Record<string, unknown> }[];
  summary: string;
}

/**
 * Multi-criteria query and count tool.
 * Enables the LLM to filter, count, and inspect spreadsheet rows matching one or more conditions
 * (e.g. Column J contains "pradeep" AND Column D contains "OUT") in a single deterministic pass.
 */
export function querySheetRecords(
  workbook: Workbook,
  sheetName: string | undefined,
  conditions: QuerySheetCondition[],
  limit = 100,
): QuerySheetRecordsResult | { error: string } {
  const sheet = resolveSheet(workbook, sheetName);
  if (!sheet) return { error: `Sheet "${sheetName ?? ''}" not found.` };

  const totalRows = sheet.rows.length;
  const boundedLimit = safeLimit(limit, 100, 500);
  const headerRow = sheet.rows[0] ?? [];
  const headers = headerRow.map((c, i) => String(c?.value ?? indexToColumn(i)));
  const sampleHeaders = recordHeaders(headerRow);
  if (!Array.isArray(conditions) || conditions.length > 50) {
    return { error: 'Supply an array of at most 50 query conditions.' };
  }

  // Resolve condition column indexes
  const resolvedConditions = [];
  for (const cond of conditions) {
    if (
      !cond ||
      (typeof cond.value !== 'string' && typeof cond.value !== 'number') ||
      (typeof cond.value === 'number' && !Number.isFinite(cond.value))
    ) {
      return { error: 'Every condition must supply a string or numeric value.' };
    }
    if (
      (cond.column !== undefined && (typeof cond.column !== 'string' || !cond.column.trim())) ||
      (cond.header !== undefined && (typeof cond.header !== 'string' || !cond.header.trim())) ||
      (cond.operator !== undefined &&
        !['equals', 'contains', 'startsWith', 'endsWith', 'gt', 'lt'].includes(cond.operator))
    ) {
      return { error: 'Condition columns, headers, and comparison operators must be valid.' };
    }
    let colIdx = cond.column ? columnToIndex(cond.column) : undefined;
    if (cond.column && (colIdx === undefined || colIdx >= sheetMaxCols(sheet.rows))) {
      return { error: `Column "${cond.column}" not found in sheet "${sheet.name}".` };
    }
    if (colIdx === undefined && cond.header) {
      const hLower = cond.header.trim().toLowerCase();
      const exact = headers.flatMap((h, i) => (h.trim().toLowerCase() === hLower ? [i] : []));
      const matches =
        exact.length > 0
          ? exact
          : headers.flatMap((h, i) => (h.toLowerCase().includes(hLower) ? [i] : []));
      if (!hLower || matches.length !== 1)
        return { error: `Header "${cond.header}" is missing or ambiguous; use a column letter.` };
      colIdx = matches[0];
    }
    const numVal = toNumericOrNull(cond.value);
    if (
      (cond.operator === 'gt' || cond.operator === 'lt') &&
      (colIdx === undefined || numVal === null)
    ) {
      return { error: 'Numeric comparisons require a valid column and numeric value.' };
    }
    resolvedConditions.push({
      ...cond,
      colIdx,
      operator: cond.operator ?? 'contains',
      strVal: String(cond.value).trim().toLowerCase(),
      numVal,
    });
  }

  const matchingRowNumbers: number[] = [];
  const sampleMatchingRows: { rowNumber: number; cells: Record<string, unknown> }[] = [];
  const read = createWorkbookValueReader(workbook);
  let totalMatchingRows = 0;

  for (let r = 1; r < sheet.rows.length; r += 1) {
    const row = sheet.rows[r] ?? [];
    let matchesAll = true;

    for (const cond of resolvedConditions) {
      let cellValStr = '';
      let cellNumVal: number | null = null;

      if (cond.colIdx !== undefined && cond.colIdx >= 0) {
        const val = read(sheet.name, cond.colIdx, r + 1);
        cellValStr = val !== null && val !== undefined ? String(val).trim().toLowerCase() : '';
        cellNumVal = toNumericOrNull(val);
      } else {
        // Search across all cells in the row if column not specified
        cellValStr = row
          .map((_cell, column) => String(read(sheet.name, column, r + 1) ?? '').toLowerCase())
          .join(' ');
      }

      let matchesCond = false;
      switch (cond.operator) {
        case 'equals':
          matchesCond = cellValStr === cond.strVal;
          break;
        case 'startsWith':
          matchesCond = cellValStr.startsWith(cond.strVal);
          break;
        case 'endsWith':
          matchesCond = cellValStr.endsWith(cond.strVal);
          break;
        case 'gt':
          matchesCond = cellNumVal !== null && cond.numVal !== null && cellNumVal > cond.numVal;
          break;
        case 'lt':
          matchesCond = cellNumVal !== null && cond.numVal !== null && cellNumVal < cond.numVal;
          break;
        case 'contains':
        default:
          matchesCond = cellValStr.includes(cond.strVal);
          break;
      }

      if (!matchesCond) {
        matchesAll = false;
        break;
      }
    }

    if (matchesAll) {
      totalMatchingRows += 1;
      if (matchingRowNumbers.length < 100) matchingRowNumbers.push(r + 1);
      if (sampleMatchingRows.length < boundedLimit) {
        const cells: Record<string, unknown> = Object.fromEntries(
          row.map((_cell, idx) => [
            sampleHeaders[idx] || indexToColumn(idx),
            read(sheet.name, idx, r + 1),
          ]),
        );
        sampleMatchingRows.push({ rowNumber: r + 1, cells });
      }
    }
  }

  const condDesc = conditions
    .map(
      (c) =>
        `${c.column ? `Column ${c.column}` : c.header || 'Row'} ${c.operator ?? 'contains'} "${c.value}"`,
    )
    .join(' AND ');

  return {
    sheet: sheet.name,
    totalMatchingRows,
    totalSheetRows: Math.max(0, totalRows - 1),
    matchingRowNumbers,
    sampleMatchingRows,
    summary: `Found ${totalMatchingRows} matching row(s) out of ${Math.max(0, totalRows - 1)} data rows in ${sheet.name} (${condDesc || 'all criteria'}).`,
  };
}

export interface WebSearchResult {
  query: string;
  results: { title: string; snippet: string }[];
  summary: string;
}

const FORMULA_KNOWLEDGE_BASE: Array<{ keywords: string[]; title: string; snippet: string }> = [
  {
    keywords: ['xlookup', 'lookup'],
    title: 'Excel XLOOKUP Formula',
    snippet:
      '=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]). Modern replacement for VLOOKUP/INDEX-MATCH that works in any direction.',
  },
  {
    keywords: ['vlookup'],
    title: 'Excel VLOOKUP Formula',
    snippet:
      '=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup]). Looks up a value in the leftmost column and returns a value in the same row from a specified column.',
  },
  {
    keywords: ['cagr', 'compound annual growth'],
    title: 'CAGR (Compound Annual Growth Rate) Formula',
    snippet:
      '=(Ending_Value/Beginning_Value)^(1/Number_of_Years) - 1. Computes constant annual rate of growth over a multi-year period.',
  },
  {
    keywords: ['yoy', 'year over year', 'growth rate'],
    title: 'Year-over-Year (YoY) Growth Formula',
    snippet:
      '=(Current_Period - Prior_Period) / Prior_Period. Express as percentage to show growth rate compared to the same period in the previous year.',
  },
  {
    keywords: ['sumifs', 'conditional sum'],
    title: 'Excel SUMIFS Formula',
    snippet:
      '=SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2, ...]). Sums cells that meet multiple criteria across columns.',
  },
  {
    keywords: ['countifs', 'conditional count'],
    title: 'Excel COUNTIFS Formula',
    snippet:
      '=COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2, ...]). Counts cells across multiple ranges that satisfy all given conditions.',
  },
  {
    keywords: ['npv', 'net present value'],
    title: 'Excel NPV Formula',
    snippet:
      '=NPV(rate, value1, [value2], ...) + Initial_Investment. Calculates the net present value of an investment using a discount rate and a series of future cash flows.',
  },
  {
    keywords: ['irr', 'internal rate of return'],
    title: 'Excel IRR Formula',
    snippet:
      '=IRR(values, [guess]). Returns the internal rate of return for a series of periodic cash flows (initial outlay as negative number).',
  },
  {
    keywords: ['margin', 'gross margin'],
    title: 'Gross Margin Percentage',
    snippet:
      '=(Revenue - COGS) / Revenue. Represents the percent of total sales revenue that the company retains after incurring the direct costs.',
  },
  {
    keywords: ['markup'],
    title: 'Markup Percentage',
    snippet:
      '=(Selling_Price - Unit_Cost) / Unit_Cost. The percentage added to the cost price of goods to cover overhead and profit.',
  },
  {
    keywords: ['stdev', 'standard deviation'],
    title: 'Excel Standard Deviation Formula',
    snippet:
      '=STDEV.S(number1, [number2], ...) for sample standard deviation; =STDEV.P(...) for entire population. Measures the dispersion of values relative to their mean.',
  },
];

/** Search the public web for Excel formula references, domain terminology, conversion rates, or facts. */
export async function searchWebKnowledge(query: string, limit = 4): Promise<WebSearchResult> {
  const needle = query.trim().toLowerCase();
  if (!needle) return { query, results: [], summary: 'No search term provided.' };
  const boundedLimit = safeLimit(limit, 4, 20);

  const matchedKnowledge: { title: string; snippet: string }[] = [];
  for (const entry of FORMULA_KNOWLEDGE_BASE) {
    if (entry.keywords.some((k) => needle.includes(k) || k.includes(needle))) {
      matchedKnowledge.push({ title: entry.title, snippet: entry.snippet });
    }
  }

  let wikiHits: { title: string; snippet: string }[] = [];
  try {
    const url = `https://en.wikipedia.org/w/api.php?action=query&list=search&srsearch=${encodeURIComponent(
      query.trim(),
    )}&format=json&origin=*`;
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 3500);
    const res = await fetch(url, { signal: controller.signal }).finally(() => clearTimeout(timer));
    if (res.ok) {
      const data = (await res.json()) as {
        query?: { search?: Array<{ title?: string; snippet?: string }> };
      };
      wikiHits = (data.query?.search ?? []).slice(0, boundedLimit).map((item) => ({
        title: item.title ?? '',
        snippet: (item.snippet ?? '')
          .replace(/<[^>]+>/g, '')
          .replace(/&quot;/g, '"')
          .replace(/&amp;/g, '&'),
      }));
    }
  } catch {
    // Network timeout or offline - rely on curated knowledge base
  }

  const combined = [...matchedKnowledge, ...wikiHits].slice(0, boundedLimit);
  return {
    query,
    results: combined,
    summary:
      combined.length > 0
        ? `Found ${combined.length} external knowledge result(s) for "${query}".`
        : `No external knowledge results found for "${query}".`,
  };
}

/** Tool definitions for LLM function calling */
export const READ_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    type: 'function',
    function: {
      name: 'search_web',
      description:
        'Search the public web and external knowledgebase for domain terminology, Excel formulas (e.g. XLOOKUP, CAGR, standard deviation), accounting standards, units, or external facts to assist with data analysis.',
      parameters: {
        type: 'object',
        properties: {
          query: { type: 'string', description: 'Search keywords, formula name, or concept' },
          limit: { type: 'number', description: 'Maximum results to return (default 4)' },
        },
        required: ['query'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'query_sheet_records',
      description:
        'Query and count rows in the spreadsheet matching one or more column conditions (e.g. column J contains "pradeep" AND column D contains "OUT"). Returns the exact total match count, matching row numbers, and sample row records without mutating the sheet.',
      parameters: {
        type: 'object',
        properties: {
          sheet: { type: 'string', description: 'Sheet name (defaults to active sheet)' },
          conditions: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                column: { type: 'string', description: 'Column letter, e.g. "D" or "J"' },
                header: { type: 'string', description: 'Optional column header name' },
                operator: {
                  type: 'string',
                  enum: ['contains', 'equals', 'startsWith', 'endsWith', 'gt', 'lt'],
                  description: 'Comparison operator (default "contains")',
                },
                value: {
                  type: 'string',
                  description: 'Value to search or match against',
                },
              },
              required: ['value'],
            },
            description: 'List of column criteria that must all match in each row',
          },
          limit: {
            type: 'number',
            description: 'Maximum sample records to return (default 25)',
          },
        },
        required: ['conditions'],
      },
    },
  },
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
          endRow: { type: 'number', description: 'Ending row number (default 50)' },
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
          limit: { type: 'number', description: 'Maximum matches to return (default 50)' },
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
