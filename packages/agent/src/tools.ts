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
