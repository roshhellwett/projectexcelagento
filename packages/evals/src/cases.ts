export interface EvalExpectation {
  /** Expected engine operation, or null when the prompt should not propose one. */
  operation: string | null;
  /** Optional partial match on the proposed args. */
  args?: Record<string, unknown>;
}

export interface EvalCase {
  id: string;
  prompt: string;
  expected: EvalExpectation;
  note?: string;
}

/**
 * Golden set for ExcelAgento's deterministic planner. Every case is reproducible
 * offline (no API key), so it doubles as a regression suite for the intent layer.
 */
export const evalCases: EvalCase[] = [
  {
    id: 'dedupe-explicit',
    prompt: 'remove duplicate rows',
    expected: { operation: 'delete_duplicates' },
  },
  {
    id: 'dedupe-slang',
    prompt: 'dedup this sheet please',
    expected: { operation: 'delete_duplicates' },
  },
  {
    id: 'trim-whitespace',
    prompt: 'trim whitespace in column B',
    expected: { operation: 'normalize_text', args: { columns: ['B'] } },
  },
  {
    id: 'uppercase-column',
    prompt: 'uppercase column B',
    expected: { operation: 'normalize_text', args: { columns: ['B'] } },
  },
  {
    id: 'format-dates-column',
    prompt: 'format dates in column C to YYYY-MM-DD',
    expected: { operation: 'format_dates', args: { column: 'C' } },
  },
  {
    id: 'sort-by-letter',
    prompt: 'sort by column A ascending',
    expected: { operation: 'sort_range', args: { column: 'A', direction: 'asc' } },
  },
  {
    id: 'sort-by-name-desc',
    prompt: 'sort by amount descending',
    expected: { operation: 'sort_range', args: { direction: 'desc' } },
  },
  {
    id: 'delete-column',
    prompt: 'delete column E',
    expected: { operation: 'delete_column' },
  },
  {
    id: 'rename-column',
    prompt: 'rename column A to Reference',
    expected: { operation: 'rename_column', args: { column: 'A' } },
  },
  {
    id: 'greeting-is-conversational',
    prompt: 'hello there',
    expected: { operation: null },
    note: 'Social turns must never mutate the workbook.',
  },
  {
    id: 'summary-is-conversational',
    prompt: 'give me a summary overview of this sheet',
    expected: { operation: null },
    note: 'Informational prompts return prose, not an action.',
  },
];
