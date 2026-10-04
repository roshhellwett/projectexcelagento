/**
 * Formula error values.
 *
 * Excel errors are values, not text. Keeping them as bare strings meant `IFERROR`
 * could not tell `#DIV/0!` from a SKU like `#12345`, and arithmetic silently
 * coerced an error into `NaN`. Detection is therefore an exact match against the
 * canonical Excel codes - never a `startsWith('#')` prefix test, which swallows
 * legitimate user text.
 */

export const FORMULA_ERROR_CODES = [
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#ERROR!',
] as const;

export type FormulaErrorCode = (typeof FORMULA_ERROR_CODES)[number];

const CODE_SET: ReadonlySet<string> = new Set(FORMULA_ERROR_CODES);

/** True only for one of Excel's canonical error codes, never for arbitrary `#`-prefixed text. */
export function isFormulaError(value: unknown): value is FormulaErrorCode {
  return typeof value === 'string' && CODE_SET.has(value);
}

/** Narrows an arbitrary value to an Excel error code, passing existing codes through unchanged. */
export function asFormulaError(value: unknown): FormulaErrorCode {
  if (isFormulaError(value)) return value;
  const text = String(value ?? '');
  const upper = text.toUpperCase();
  return CODE_SET.has(upper) ? (upper as FormulaErrorCode) : '#ERROR!';
}
