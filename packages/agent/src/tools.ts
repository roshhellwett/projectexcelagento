import { z } from 'zod';

import type { OperationRegistry } from '@excel-agent/engine';

import type { ProposedAction } from './analysis.js';

export interface ToolDescriptor {
  name: string;
  description: string;
  category: ProposedAction['category'];
  /** JSON Schema derived live from the engine's Zod contract. */
  jsonSchema: Record<string, unknown>;
  example: Record<string, unknown>;
}

interface ToolMetadata {
  description: string;
  category: ProposedAction['category'];
  example: Record<string, unknown>;
}

/**
 * Human-readable metadata per engine operation. The set of names is asserted to
 * match the live registry, so adding an engine operation fails loudly here until
 * it is documented for the model.
 */
const TOOL_METADATA: Readonly<Record<string, ToolMetadata>> = {
  format_cells: {
    description:
      'Apply workbook-native formatting to a rectangular cell range: bold, italic, underline, colors, alignment, wrapping, and number formats.',
    category: 'format',
    example: {
      sheet: 'Sheet1',
      startRow: 1,
      endRow: 10,
      startColumn: 'A',
      endColumn: 'D',
      style: { bold: true, fillColor: '#EAF2E6' },
    },
  },
  format_dates: {
    description: 'Normalize a column of dates into a single, consistent format.',
    category: 'format',
    example: { sheet: 'Sheet1', column: 'C', format: 'YYYY-MM-DD', headerRow: 1 },
  },
  normalize_text: {
    description:
      'Trim, collapse whitespace, and optionally change case (upper/lower/title) across one or more text columns.',
    category: 'transform',
    example: {
      sheet: 'Sheet1',
      columns: ['B'],
      trim: true,
      collapseWhitespace: true,
      case: 'title',
      headerRow: 1,
    },
  },
  sort_range: {
    description: 'Sort the data rows of a sheet by a column, ascending or descending.',
    category: 'transform',
    example: { sheet: 'Sheet1', column: 'A', direction: 'asc', startRow: 2, startColumn: 'A' },
  },
  filter_rows: {
    description:
      'Keep only the rows matching a column condition (equals, contains, gt, lt, is_blank, ...).',
    category: 'filter',
    example: {
      sheet: 'Sheet1',
      column: 'D',
      operator: 'equals',
      value: 'Completed',
      headerRow: 1,
    },
  },
  find_replace: {
    description: 'Find and replace text across the sheet, optionally case-sensitive or whole-cell.',
    category: 'transform',
    example: {
      sheet: 'Sheet1',
      find: 'Acme',
      replace: 'ACME Corp',
      matchCase: false,
      wholeCell: false,
    },
  },
  delete_duplicates: {
    description: 'Remove duplicate rows, keyed on the given columns, keeping the first occurrence.',
    category: 'transform',
    example: { sheet: 'Sheet1', columns: ['A'], keep: 'first', headerRow: 1 },
  },
  rename_column: {
    description: 'Rename a column header.',
    category: 'columns',
    example: { sheet: 'Sheet1', column: 'A', newName: 'Order Reference', headerRow: 1 },
  },
  delete_column: {
    description: 'Delete an entire column from the sheet.',
    category: 'structure',
    example: { sheet: 'Sheet1', column: 'F' },
  },
  add_column: {
    description: 'Append a new column with a header and an optional default value.',
    category: 'structure',
    example: {
      sheet: 'Sheet1',
      column: 'Z',
      headerName: 'Tax Rate',
      defaultValue: 0,
      headerRow: 1,
    },
  },
  set_cells: {
    description: 'Set explicit values for one or more individual cells.',
    category: 'transform',
    example: { sheet: 'Sheet1', cells: [{ column: 'A', row: 2, value: 'New value' }] },
  },
  fill_blanks: {
    description:
      'Fill empty or blank cells in a column using a strategy (forward fill, backward fill, mean, or static value).',
    category: 'transform',
    example: { sheet: 'Sheet1', column: 'B', strategy: 'forward', headerRow: 1 },
  },
  add_computed_column: {
    description:
      'Create a new calculated column by evaluating an arithmetic or string expression across existing columns.',
    category: 'columns',
    example: {
      sheet: 'Sheet1',
      headerName: 'Total',
      expression: "col('Price') * col('Quantity')",
      headerRow: 1,
    },
  },
  split_column: {
    description: 'Split a column into two or more new columns by a delimiter string.',
    category: 'columns',
    example: {
      sheet: 'Sheet1',
      column: 'B',
      delimiter: ' ',
      newColumnNames: ['First Name', 'Last Name'],
      headerRow: 1,
    },
  },
  merge_columns: {
    description: 'Combine multiple columns into a single column joined by a separator string.',
    category: 'columns',
    example: {
      sheet: 'Sheet1',
      columns: ['A', 'B'],
      separator: ' ',
      headerName: 'Full Name',
      headerRow: 1,
    },
  },
  clean_to_new_sheet: {
    description:
      'Produce a cleaned, structured copy of the data into a new sheet: trim whitespace, drop blank/banner rows and empty columns, auto-detect the header row, and coerce numbers.',
    category: 'transform',
    example: { sheet: 'Sheet1', targetSheet: 'Sheet1_Clean' },
  },
  edit_cells: {
    description:
      'Write specific values into individual cells, growing the sheet when an address is past the current data. Use for precise, targeted edits; prefer the range operations for bulk work.',
    category: 'transform',
    example: { sheet: 'Sheet1', edits: [{ row: 2, column: 'B', value: 60 }] },
  },
  filter_to_new_sheet: {
    description:
      'Filter rows matching a condition (equals, contains, starts_with, gt, lt, is_blank, etc.) and extract or copy them into a brand new worksheet with headers. Use this whenever the user asks to filter/extract/separate data into a new, separate, or different sheet.',
    category: 'filter',
    example: {
      sheet: 'Sheet1',
      targetSheet: 'IN_Data',
      column: 'D',
      operator: 'contains',
      value: 'IN',
      headerRow: 1,
    },
  },
  create_sheet: {
    description:
      'Create a new empty or titled worksheet with optional column headers and optional initial data rows.',
    category: 'structure',
    example: {
      sheetName: 'Summary',
      headers: ['ID', 'Date', 'Amount'],
      rows: [['1', '2024-01-01', 100]],
    },
  },
  append_rows: {
    description:
      'Append multiple new data rows to an existing worksheet. Use this to insert extracted, migrated, or newly generated records into a target table.',
    category: 'transform',
    example: {
      sheet: 'Sheet1',
      rows: [
        ['Acme Corp', 'Finance', '123-456-7890'],
        ['Beta LLC', 'Tech', '987-654-3210'],
      ],
    },
  },
  duplicate_sheet: {
    description: 'Duplicate an entire existing worksheet into a new sheet.',
    category: 'structure',
    example: { sheet: 'Sheet1', targetSheet: 'Sheet1_Backup' },
  },
  delete_sheet: {
    description: 'Delete an entire worksheet from the workbook. Demands confirmation.',
    category: 'structure',
    example: { sheet: 'OldSheet' },
  },
  add_summary_row: {
    description:
      'Append a summary row (Total, Average, Count, Min, Max) at the bottom of the data for numeric columns.',
    category: 'transform',
    example: {
      sheet: 'Sheet1',
      aggregation: 'sum',
      label: 'Total',
      columns: ['E', 'H'],
      headerRow: 1,
    },
  },
  lookup_merge: {
    description:
      'Merge and match values from another sheet based on a common key column (VLOOKUP / XLOOKUP behavior).',
    category: 'columns',
    example: {
      sheet: 'Sheet1',
      keyColumn: 'B',
      lookupSheet: 'Customers',
      lookupKeyColumn: 'A',
      lookupValueColumn: 'B',
      headerName: 'Customer Name',
      headerRow: 1,
    },
  },
  aggregate_column: {
    description:
      'Compute one number for a column: sum, average, count, count_distinct, min, max, median, or stdev. Writes nothing and returns the value in the report. criteria takes a COUNTIF expression such as ">100", or an operator with its operand in criteriaValue.',
    category: 'transform',
    example: {
      sheet: 'Sheet1',
      column: 'D',
      aggregation: 'sum',
      criteria: '>100',
    },
  },
  group_and_summarize: {
    description:
      'Pivot: group rows by one or more columns, aggregate a value column (sum, average, count, count_distinct, min, max, median, stdev), and write the result to a new sheet, sorted by group key. Text in the value column is counted and reported, not read as zero.',
    category: 'transform',
    example: {
      sheet: 'Sheet1',
      groupBy: ['B'],
      valueColumn: 'D',
      aggregation: 'sum',
      targetSheet: 'By Region',
    },
  },
  join_sheets: {
    description:
      'Join two sheets on a key column, appending one or more lookup columns. Keys match case-insensitively after trimming; duplicate lookup keys keep the first match and warn. An inner join drops unmatched rows and needs confirmation; a left join keeps them blank.',
    category: 'columns',
    example: {
      sheet: 'Sheet1',
      keyColumn: 'B',
      lookupSheet: 'Customers',
      lookupKeyColumn: 'A',
      lookupValueColumn: 'B',
      headerName: 'Customer Name',
      joinType: 'left',
    },
  },
  fill_series: {
    description:
      'Autofill a column below a source range: copy repeats values, linear extrapolates a step, date repeats a day/week/month/year interval, text increments a trailing number. Refuses when the source cannot support the strategy and confirms before overwriting.',
    category: 'transform',
    example: {
      sheet: 'Sheet1',
      column: 'A',
      strategy: 'linear',
      sourceRange: 'A2:A3',
      targetStartRow: 4,
      targetEndRow: 12,
    },
  },
  categorize_column: {
    description:
      'Label each row with the first matching rule (equals, not_equals, contains, starts_with, ends_with, gt/gte/lt/lte, between) or otherwise. Refuses a duplicate column name unless afterColumn places it explicitly.',
    category: 'columns',
    example: {
      sheet: 'Sheet1',
      sourceColumn: 'D',
      newColumnName: 'Amount Band',
      rules: [
        { operator: 'gte', value: 1000, label: 'Large' },
        { operator: 'between', value: [100, 999], label: 'Medium' },
      ],
      otherwise: 'Small',
    },
  },
  reconcile_sheets: {
    description:
      'Reconcile two sheets on matching key columns (e.g. Orders vs Invoices) and compare amounts. Generates four comprehensive report sheets: Summary, Matched records, Exceptions/Discrepancies, and Methodology, while keeping source sheets completely untouched.',
    category: 'transform',
    example: {
      leftSheet: 'Orders',
      rightSheet: 'Invoices',
      leftKeys: ['A'],
      rightKeys: ['A'],
      leftAmount: 'C',
      rightAmount: 'C',
      tolerance: 0,
      reportPrefix: 'Reconciliation',
    },
  },
};

export class ToolCatalogError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ToolCatalogError';
  }
}

/**
 * Build the tool catalog from the live engine registry. JSON schemas are derived
 * from each operation's Zod schema so the prompt can never drift from the engine.
 */
export function buildToolCatalog(
  registry: Pick<OperationRegistry, 'names' | 'getSchema'>,
): ToolDescriptor[] {
  const names = [...registry.names];
  const undocumented = names.filter((name) => !(name in TOOL_METADATA));
  if (undocumented.length > 0) {
    throw new ToolCatalogError(
      `Engine operations without tool metadata: ${undocumented.join(', ')}.`,
    );
  }

  return names.map((name) => {
    const metadata = TOOL_METADATA[name] as ToolMetadata;
    const schema = registry.getSchema(name);
    let jsonSchema: Record<string, unknown> = {};
    try {
      jsonSchema = schema
        ? (z.toJSONSchema(schema, { io: 'input', unrepresentable: 'any' }) as Record<
            string,
            unknown
          >)
        : {};
    } catch {
      jsonSchema = {};
    }
    return {
      name,
      description: metadata.description,
      category: metadata.category,
      jsonSchema,
      example: metadata.example,
    };
  });
}

/** Render the catalog as a compact, model-friendly contract block. */
export function describeTools(catalog: ToolDescriptor[]): string {
  return catalog
    .map((tool) => {
      const example = JSON.stringify({ name: tool.name, args: tool.example });
      const schemaObj: Record<string, unknown> = {
        type: 'object',
        properties: (tool.jsonSchema as Record<string, unknown>).properties ?? {},
      };
      if ((tool.jsonSchema as Record<string, unknown>).required) {
        schemaObj.required = (tool.jsonSchema as Record<string, unknown>).required;
      }
      return `- ${tool.name} [${tool.category}]: ${tool.description}\n  schema: ${JSON.stringify(
        schemaObj,
      )}\n  example: ${example}`;
    })
    .join('\n');
}
