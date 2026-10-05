import {
  columnToIndex,
  fitLinearRegression,
  indexToColumn,
  maxColumnCount,
  summarizeNumericValues,
  toNumericOrNull,
  createWorkbookValueReader,
  type Workbook,
  type Sheet,
} from '@excel-agent/engine';
import type { ToolDefinition } from './types.js';

function resolveColumn(
  workbook: Workbook,
  sheetName: string,
  column: string,
  headerRow: number,
): { error: string } | { sheet: Sheet; index: number; headerName: string } {
  const sheet = workbook.sheets.find(
    (sheet) => sheet.name.toLowerCase() === sheetName.trim().toLowerCase(),
  );
  if (!sheet) return { error: `Sheet "${sheetName}" not found.` };
  const index = columnToIndex(column.trim());
  if (index === undefined || index >= maxColumnCount(sheet.rows)) {
    return { error: `Column "${column}" not found.` };
  }
  if (!Number.isInteger(headerRow) || headerRow < 1 || headerRow > sheet.rows.length) {
    return { error: 'headerRow must identify an existing row.' };
  }
  return {
    sheet,
    index,
    headerName: String(sheet.rows[headerRow - 1]?.[index]?.value ?? indexToColumn(index)),
  };
}

/** Full-column statistics, never a calculation over the bounded prompt sample. */
export function describeColumn(
  workbook: Workbook,
  sheetName: string,
  column: string,
  headerRow = 1,
) {
  const resolved = resolveColumn(workbook, sheetName, column, headerRow);
  if ('error' in resolved) return resolved;
  const { sheet, index, headerName } = resolved;
  const read = createWorkbookValueReader(workbook);
  const values = sheet.rows
    .slice(headerRow)
    .map((_row, offset) => read(sheet.name, index, headerRow + offset + 1));
  const summary = summarizeNumericValues(values);
  const outliers: { rowNumber: number; value: number }[] = [];
  if (summary.lowerFence !== null && summary.upperFence !== null) {
    for (const [offset, value] of values.entries()) {
      const numeric = toNumericOrNull(value);
      if (
        numeric !== null &&
        (numeric < summary.lowerFence || numeric > summary.upperFence) &&
        outliers.length < 25
      ) {
        outliers.push({ rowNumber: headerRow + offset + 1, value: numeric });
      }
    }
  }
  return {
    sheet: sheet.name,
    column: indexToColumn(index),
    headerName,
    headerRow,
    ...summary,
    outliers,
    outlierMethod: 'Tukey 1.5 × IQR; inclusive quartiles',
  };
}

export function analyzeColumnRelationship(
  workbook: Workbook,
  sheetName: string,
  xColumn: string,
  yColumn: string,
  headerRow = 1,
) {
  const x = resolveColumn(workbook, sheetName, xColumn, headerRow);
  if ('error' in x) return x;
  const y = resolveColumn(workbook, sheetName, yColumn, headerRow);
  if ('error' in y) return y;
  if (x.index === y.index) return { error: 'Choose two different columns.' };
  const read = createWorkbookValueReader(workbook);
  const pairs = x.sheet.rows.slice(headerRow).map(
    (_row, offset) =>
      [
        read(x.sheet.name, x.index, headerRow + offset + 1),
        read(x.sheet.name, y.index, headerRow + offset + 1),
      ] as const,
  );
  return {
    sheet: x.sheet.name,
    xColumn: indexToColumn(x.index),
    yColumn: indexToColumn(y.index),
    xHeader: x.headerName,
    yHeader: y.headerName,
    headerRow,
    ...fitLinearRegression(pairs),
    method: 'Pearson correlation and ordinary least squares with an intercept; complete pairs only',
  };
}

export const STATISTICAL_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    type: 'function',
    function: {
      name: 'describe_column',
      description:
        'Compute full-column descriptive statistics: counts, missing and nonnumeric exclusions, mean, median, inclusive quartiles, sample variance and standard deviation, and 1.5-IQR outliers with row numbers. Read-only.',
      parameters: {
        type: 'object',
        properties: {
          sheet: { type: 'string' },
          column: { type: 'string', description: 'Column letter' },
          headerRow: { type: 'integer', minimum: 1, description: 'Header row, default 1' },
        },
        required: ['column'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'analyze_column_relationship',
      description:
        'Compute Pearson correlation and linear regression of y on x, including slope, intercept, R², residual standard error and excluded pair count. Uses complete numeric rows; returns undefined statistics for constant or insufficient data. Read-only.',
      parameters: {
        type: 'object',
        properties: {
          sheet: { type: 'string' },
          xColumn: { type: 'string', description: 'Predictor column letter' },
          yColumn: { type: 'string', description: 'Response column letter' },
          headerRow: { type: 'integer', minimum: 1 },
        },
        required: ['xColumn', 'yColumn'],
      },
    },
  },
];
