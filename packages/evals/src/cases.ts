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

  // --- Duplicate removal variants -------------------------------------------------
  {
    id: 'dedupe-hinglish',
    prompt: 'duplicate rows hatao',
    expected: { operation: 'delete_duplicates' },
    note: 'Hinglish phrasing still routes to the dedupe intent.',
  },
  {
    id: 'dedupe-typo-tolerated',
    prompt: 'remove dublicate rows',
    expected: { operation: 'delete_duplicates' },
    note: 'Common misspelling "dublicate" is explicitly recognised.',
  },
  {
    id: 'dedupe-unique-phrasing',
    prompt: 'keep only unique rows',
    expected: { operation: 'delete_duplicates' },
    note: 'The "unique" synonym must not be confused with a unique-values request.',
  },

  // --- Text normalization --------------------------------------------------------
  {
    id: 'normalize-lowercase-by-header-name',
    prompt: 'lowercase the customer column',
    expected: { operation: 'normalize_text', args: { columns: ['B'] } },
    note: 'Column resolved from the header name rather than a letter.',
  },
  {
    id: 'normalize-titlecase-by-header-name',
    prompt: 'title case the status column',
    expected: { operation: 'normalize_text', args: { columns: ['E'] } },
    note: 'Title case is a distinct casing mode, not a trim.',
  },
  {
    id: 'normalize-trim-hinglish',
    prompt: 'column B ko trim karo',
    expected: { operation: 'normalize_text', args: { columns: ['B'] } },
    note: 'Hinglish trim request resolves the letter-named column.',
  },
  {
    id: 'normalize-caps-by-header-name',
    prompt: 'convert order id to caps',
    expected: { operation: 'normalize_text', args: { columns: ['A'] } },
    note: '"caps" is an uppercase synonym; multi-word headers must resolve.',
  },

  // --- Date normalization --------------------------------------------------------
  {
    id: 'dates-iso-default',
    prompt: 'normalize the order date to YYYY-MM-DD',
    expected: { operation: 'format_dates', args: { column: 'C', format: 'YYYY-MM-DD' } },
    note: 'ISO is the default when no locale is specified.',
  },
  {
    id: 'dates-us-format',
    prompt: 'convert the order date to US format',
    expected: { operation: 'format_dates', args: { column: 'C', format: 'MM/DD/YYYY' } },
    note: 'A standalone "us" token selects the US date format.',
  },
  {
    id: 'dates-eu-format',
    prompt: 'show order date in EU format',
    expected: { operation: 'format_dates', args: { column: 'C', format: 'DD/MM/YYYY' } },
    note: 'A standalone "eu" token selects the European date format.',
  },
  {
    id: 'dates-us-substring-is-not-a-locale',
    prompt: 'the order date must be YYYY-MM-DD',
    expected: { operation: 'format_dates', args: { column: 'C', format: 'YYYY-MM-DD' } },
    note: 'Regression: "must" contains the substring "us" and must not imply MM/DD/YYYY.',
  },
  {
    id: 'dates-iso-hinglish',
    prompt: 'column C ke dates ko YYYY-MM-DD format me karo',
    expected: { operation: 'format_dates', args: { column: 'C', format: 'YYYY-MM-DD' } },
    note: 'Hinglish date formatting with an explicit target format.',
  },

  // --- Sorting -------------------------------------------------------------------
  {
    id: 'sort-highest-amount-first',
    prompt: 'sort by highest amount first',
    expected: { operation: 'sort_range', args: { column: 'D', direction: 'desc' } },
    note: '"highest" implies descending without the word "descending".',
  },
  {
    id: 'sort-order-by-z-a',
    prompt: 'order by status z-a',
    expected: { operation: 'sort_range', args: { column: 'E', direction: 'desc' } },
    note: '"order by" and "z-a" are both supported synonyms.',
  },
  {
    id: 'sort-hinglish-descending',
    prompt: 'amount ko descending order me sort karo',
    expected: { operation: 'sort_range', args: { column: 'D', direction: 'desc' } },
    note: 'Hinglish sort request resolved by header name.',
  },

  // --- Column structural ---------------------------------------------------------
  {
    id: 'delete-column-by-header-name',
    prompt: 'drop column Amount',
    expected: { operation: 'delete_column', args: { column: 'D' } },
    note: '"drop" is a delete synonym and the header name must resolve.',
  },
  {
    id: 'rename-column-preserves-header-case',
    prompt: 'rename column C to Order Date ISO',
    expected: { operation: 'rename_column', args: { column: 'C', newName: 'Order Date ISO' } },
    note: 'Regression: the new header must keep the casing the user typed.',
  },

  // --- Find & replace ------------------------------------------------------------
  {
    id: 'find-replace-preserves-case',
    prompt: 'replace Pending with Completed',
    expected: { operation: 'find_replace', args: { find: 'Pending', replace: 'Completed' } },
    note: 'Regression: replacement text must keep its casing, not be lowercased.',
  },
  {
    id: 'find-and-replace-phrasing',
    prompt: 'find Acme and replace with Acme Corporation',
    expected: {
      operation: 'find_replace',
      args: { find: 'Acme', replace: 'Acme Corporation' },
    },
    note: 'Verbose "find X and replace with Y" phrasing.',
  },

  // --- Value-based filtering -----------------------------------------------------
  {
    id: 'filter-numeric-greater-than',
    prompt: 'filter amount greater than 1000',
    expected: { operation: 'filter_rows', args: { column: 'D', operator: 'gt', value: 1000 } },
    note: 'Comparison operators must carry the numeric threshold through.',
  },
  {
    id: 'filter-numeric-less-than-hinglish',
    prompt: 'amount kam 1000 wale rows dikhao',
    expected: { operation: 'filter_rows', args: { column: 'D', operator: 'lt', value: 1000 } },
    note: 'Hinglish "kam" maps to a less-than comparison.',
  },
  {
    id: 'filter-categorical-equals',
    prompt: 'filter status is completed',
    expected: {
      operation: 'filter_rows',
      args: { column: 'E', operator: 'equals', value: 'Completed' },
    },
    note: 'Categorical matching against a distinct value in the column.',
  },

  // --- Informational turns: prose, never a mutation --------------------------------
  {
    id: 'math-sum-is-informational',
    prompt: 'what is the total amount',
    expected: { operation: null },
    note: 'Aggregations are answered in prose; they are not engine operations.',
  },
  {
    id: 'math-average-is-informational',
    prompt: 'average of the amount column',
    expected: { operation: null },
    note: 'Statistical summaries stay read-only.',
  },
  {
    id: 'pivot-top-category-is-informational',
    prompt: 'who handled the most transactions',
    expected: { operation: null },
    note: 'Distribution analysis reports findings without mutating data.',
  },
  {
    id: 'pivot-hinglish-is-informational',
    prompt: 'sabse zyada kaun order karta hai',
    expected: { operation: null },
    note: 'Hinglish pivot question stays informational.',
  },
  {
    id: 'missing-values-audit',
    prompt: 'find missing values',
    expected: { operation: null },
    note: 'A missing-value audit reports counts but proposes no operation.',
  },
  {
    id: 'missing-values-hinglish',
    prompt: 'khali cells dikhao',
    expected: { operation: null },
    note: 'Hinglish "khali" (empty) must not trigger a mutation.',
  },
  {
    id: 'summary-hinglish-is-conversational',
    prompt: 'is sheet ka kya hai',
    expected: { operation: null },
    note: 'Hinglish summary question returns an overview, not an action.',
  },
  {
    id: 'casual-farewell-no-action',
    prompt: 'thanks, talk later',
    expected: { operation: null },
    note: 'Politeness must never be misread as an instruction.',
  },
];
