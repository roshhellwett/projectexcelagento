import { FORMULA_FUNCTIONS, getFormulaDateSystem, setFormulaDateSystem } from './functions.js';
import { FORMULA_ERROR_CODES, isFormulaError, type FormulaErrorCode } from './errors.js';
import { daysBetween } from './excel-date.js';
import { columnToIndex } from '../workbook.js';
import type { FormulaContext, FormulaValue } from './types.js';

type ParsedValue = FormulaValue | FormulaValue[][];

/**
 * Guards against runaway recursion from self-referential formulas. Excel raises a
 * circular-reference error rather than hanging; without this the engine died by stack
 * exhaustion and reported a misleading `#ERROR!`.
 */
const MAX_EVALUATION_DEPTH = 64;

interface Token {
  type:
    | 'NUMBER'
    | 'STRING'
    | 'BOOLEAN'
    | 'IDENT'
    | 'CELL'
    | 'RANGE'
    | 'OP'
    | 'LPAREN'
    | 'RPAREN'
    | 'COMMA'
    | 'ERROR'
    | 'INVALID';
  value: string;
  sheet?: string;
}

const CELL_REGEX = /^(?:([A-Za-z0-9_]+)!)?\$?([A-Za-z]+)\$?(\d+)$/;
const RANGE_REGEX = /^(?:([A-Za-z0-9_]+)!)?\$?([A-Za-z]+)\$?(\d+):\$?([A-Za-z]+)\$?(\d+)$/;
const MAX_FORMULA_LENGTH = 8_192;

function validReference(column: string, row: number): boolean {
  const index = columnToIndex(column);
  return (
    index !== undefined &&
    index <= 16_383 &&
    Number.isSafeInteger(row) &&
    row >= 1 &&
    row <= 1_048_576
  );
}

export function tokenize(formulaStr: string): Token[] {
  let str = formulaStr.trim();
  if (str.startsWith('=')) str = str.slice(1).trim();
  if (str.length > MAX_FORMULA_LENGTH) {
    return [{ type: 'INVALID', value: 'formula exceeds Excel length limit' }];
  }

  const tokens: Token[] = [];
  let i = 0;

  while (i < str.length) {
    const ch = str[i]!;

    // Whitespace
    if (/\s/.test(ch)) {
      i++;
      continue;
    }

    // String literals. A single-quoted token immediately followed by '!' is a quoted sheet name.
    if (ch === '"' || ch === "'") {
      const quote = ch;
      let text = '';
      let j = i + 1;
      while (j < str.length) {
        if (str[j] === quote) {
          if (str[j + 1] === quote) {
            text += quote;
            j += 2;
            continue;
          }
          break;
        }
        text += str[j];
        j++;
      }
      if (j >= str.length) {
        tokens.push({ type: 'INVALID', value: 'unterminated string' });
        break;
      }
      if (quote === "'" && str[j + 1] === '!') {
        // Quoted sheet reference, e.g. 'My Sheet'!A1:B2
        let k = j + 2;
        while (k < str.length && /[A-Za-z0-9_$!:]/.test(str[k]!)) k++;
        const ref = str.slice(j + 2, k);
        const sheet = text;
        const rangeMatch = ref.match(/^\$?([A-Za-z]+)\$?(\d+):\$?([A-Za-z]+)\$?(\d+)$/);
        const cellMatch = ref.match(/^\$?([A-Za-z]+)\$?(\d+)$/);
        if (rangeMatch) {
          tokens.push({ type: 'RANGE', value: `${sheet}!${ref}`, sheet });
        } else if (cellMatch) {
          tokens.push({ type: 'CELL', value: `${sheet}!${ref}`, sheet });
        } else {
          tokens.push({ type: 'INVALID', value: ref });
        }
        i = k;
        continue;
      }
      i = j + 1; // closing quote
      tokens.push({ type: 'STRING', value: text });
      continue;
    }

    if (ch === '#') {
      const error = FORMULA_ERROR_CODES.find(
        (code) => str.slice(i, i + code.length).toUpperCase() === code,
      );
      if (error) {
        tokens.push({ type: 'ERROR', value: error });
        i += error.length;
        continue;
      }
    }

    // Numbers (incl. scientific notation 1E5, 1.5e-3)
    if (/[0-9]/.test(ch) || (ch === '.' && /[0-9]/.test(str[i + 1] ?? ''))) {
      let numStr = '';
      while (i < str.length && /[0-9.]/.test(str[i]!)) {
        numStr += str[i];
        i++;
      }
      if ((str[i] === 'e' || str[i] === 'E') && /[0-9+-]/.test(str[i + 1] ?? '')) {
        numStr += str[i++]!;
        if (str[i] === '+' || str[i] === '-') numStr += str[i++]!;
        while (i < str.length && /[0-9]/.test(str[i]!)) numStr += str[i++]!;
      }
      tokens.push({ type: 'NUMBER', value: numStr });
      continue;
    }

    // Two-character operators: <=, >=, <>, !=, ==
    const two = str.slice(i, i + 2);
    if (['<=', '>=', '<>', '!=', '=='].includes(two)) {
      tokens.push({ type: 'OP', value: two === '!=' ? '<>' : two === '==' ? '=' : two });
      i += 2;
      continue;
    }

    // Single-character operators & syntax
    if (['+', '-', '*', '/', '^', '&', '=', '<', '>', '%'].includes(ch)) {
      tokens.push({ type: 'OP', value: ch });
      i++;
      continue;
    }

    if (ch === '(') {
      tokens.push({ type: 'LPAREN', value: '(' });
      i++;
      continue;
    }
    if (ch === ')') {
      tokens.push({ type: 'RPAREN', value: ')' });
      i++;
      continue;
    }
    if (ch === ',') {
      tokens.push({ type: 'COMMA', value: ',' });
      i++;
      continue;
    }

    // Identifiers, Cell refs, Range refs, Booleans
    let word = '';
    while (i < str.length && /[A-Za-z0-9_!:$]/.test(str[i]!)) {
      word += str[i];
      i++;
    }

    if (word === '') {
      // Unsupported syntax must not disappear into a plausible partial result.
      tokens.push({ type: 'INVALID', value: ch });
      i++;
      continue;
    }

    if (word.toUpperCase() === 'TRUE' || word.toUpperCase() === 'FALSE') {
      tokens.push({ type: 'BOOLEAN', value: word.toUpperCase() });
      continue;
    }

    // Check if range: Sheet1!A1:B10 or A1:B10
    const rangeMatch = word.match(RANGE_REGEX);
    if (rangeMatch) {
      tokens.push({ type: 'RANGE', value: word, sheet: rangeMatch[1] });
      continue;
    }

    // Check if cell: Sheet1!A1 or A1
    const cellMatch = word.match(CELL_REGEX);
    if (cellMatch) {
      tokens.push({ type: 'CELL', value: word, sheet: cellMatch[1] });
      continue;
    }

    // Otherwise, function name or identifier
    tokens.push({ type: 'IDENT', value: word.toUpperCase() });
  }

  return tokens;
}

function firstError(...values: ParsedValue[]): FormulaErrorCode | null {
  for (const value of values) {
    if (isFormulaError(value)) return value;
  }
  return null;
}

/** Excel's numeric coercion: booleans are 1/0, numeric text is parsed, anything else is `#VALUE!`. */
function toNumeric(value: ParsedValue): number | FormulaErrorCode {
  if (value === null) return 0;
  if (typeof value === 'number') return Number.isFinite(value) ? value : '#NUM!';
  if (typeof value === 'boolean') return value ? 1 : 0;
  if (typeof value === 'string') {
    if (isFormulaError(value)) return value;
    const trimmed = value.trim();
    if (trimmed === '') return 0;
    const parsed = Number(trimmed.replace(/,/g, ''));
    return Number.isFinite(parsed) ? parsed : '#VALUE!';
  }
  return '#VALUE!';
}

function isDateLike(value: ParsedValue): value is Date {
  return value instanceof Date;
}

/**
 * Excel's `+`/`-`: subtracting two dates yields a day count, and adding a number to a date
 * shifts the date. Every other date combination is `#VALUE!`.
 */
function addOrSubtract(left: ParsedValue, right: ParsedValue, subtract: boolean): ParsedValue {
  const leftIsDate = isDateLike(left);
  const rightIsDate = isDateLike(right);

  if (leftIsDate && rightIsDate) {
    return subtract ? daysBetween(left, right) : '#VALUE!';
  }

  if (leftIsDate || rightIsDate) {
    const date = (leftIsDate ? left : right) as Date;
    const offset = toNumeric((leftIsDate ? right : left) as ParsedValue);
    if (isFormulaError(offset)) return offset;
    const shifted = new Date(date.getTime() + (subtract ? -offset : offset) * 86_400_000);
    return Number.isFinite(shifted.getTime()) ? shifted : '#NUM!';
  }

  const a = toNumeric(left);
  if (isFormulaError(a)) return a;
  const b = toNumeric(right);
  if (isFormulaError(b)) return b;
  const result = subtract ? a - b : a + b;
  return Number.isFinite(result) ? result : '#NUM!';
}

/** Excel's `*`, `/`, `^`: dates have no meaning here, and non-finite results are `#NUM!`. */
function arithmetic(left: ParsedValue, right: ParsedValue, op: '*' | '/' | '^'): ParsedValue {
  const a = toNumeric(left);
  if (isFormulaError(a)) return a;
  const b = toNumeric(right);
  if (isFormulaError(b)) return b;

  if (op === '/') {
    if (b === 0) return '#DIV/0!';
    const result = a / b;
    return Number.isFinite(result) ? result : '#NUM!';
  }
  const result = op === '*' ? a * b : Math.pow(a, b);
  return Number.isFinite(result) ? result : '#NUM!';
}

/** Applies an ordering predicate, letting an error operand propagate instead of comparing. */
function orderedComparison(
  left: ParsedValue,
  right: ParsedValue,
  predicate: (ordering: number) => boolean,
): ParsedValue {
  const ordering = compareOrdered(left, right);
  return isFormulaError(ordering) ? ordering : predicate(ordering);
}

function formulaEquals(a: ParsedValue, b: ParsedValue): ParsedValue {
  const propagated = firstError(a, b);
  if (propagated) return propagated;
  if (a === null && b === null) return true;
  if (a === null || b === null) return b === '' || a === '';
  if (typeof a === 'number' && typeof b === 'number') return a === b;
  if (typeof a === 'boolean' || typeof b === 'boolean') return a === b;
  // Dates compare by day, matching Excel, so a time-of-day component cannot flip equality.
  if (isDateLike(a) && isDateLike(b)) return daysBetween(a, b) === 0;
  if (isDateLike(a) || isDateLike(b)) {
    const date = (isDateLike(a) ? a : b) as Date;
    const otherNumber = toNumeric((isDateLike(a) ? b : a) as ParsedValue);
    if (isFormulaError(otherNumber)) return false;
    return date.getTime() === otherNumber;
  }
  const an = typeof a === 'string' && a.trim() !== '' ? Number(a) : NaN;
  const bn = typeof b === 'string' && b.trim() !== '' ? Number(b) : NaN;
  if (!isNaN(an) && !isNaN(bn)) return an === bn;
  if (typeof a === 'number' && !isNaN(bn as number)) return a === bn;
  if (typeof b === 'number' && !isNaN(an as number)) return b === an;
  return String(a) === String(b);
}

function compareOrdered(a: ParsedValue, b: ParsedValue): number | FormulaErrorCode {
  const propagated = firstError(a, b);
  if (propagated) return propagated;
  const rank = (v: ParsedValue): number =>
    typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 3;
  const asNumber = (v: ParsedValue): number =>
    typeof v === 'number'
      ? v
      : typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v))
        ? Number(v)
        : NaN;
  const aNum = asNumber(a);
  const bNum = asNumber(b);
  if (!isNaN(aNum) && !isNaN(bNum)) return aNum < bNum ? -1 : aNum > bNum ? 1 : 0;
  const ra = rank(a);
  const rb = rank(b);
  if (ra !== rb) return ra < rb ? -1 : 1;
  if (typeof a === 'boolean' && typeof b === 'boolean') return a === b ? 0 : a ? 1 : -1;
  const sa = String(a);
  const sb = String(b);
  return sa < sb ? -1 : sa > sb ? 1 : 0;
}

export function evaluateFormula(formula: string, context: FormulaContext): FormulaValue {
  const previousDateSystem = getFormulaDateSystem();
  setFormulaDateSystem(context.dateSystem ?? '1900');
  try {
    const tokens = tokenize(formula);
    if (tokens.length === 0) return null;

    let cursor = 0;
    let depth = 0;
    let malformed = tokens.some((token) => token.type === 'INVALID');

    /** Reports whether a sheet named in a reference actually exists, so typos become `#REF!`. */
    function sheetExists(sheet: string): boolean {
      if (!context.hasSheet) return true;
      return context.hasSheet(sheet);
    }

    function parseExpression(): ParsedValue {
      return parseLogicalOr();
    }

    function parseLogicalOr(): ParsedValue {
      const left = parseComparison();
      // In Excel, OR is usually a function, but handle infix comparison / logical
      return left;
    }

    function parseComparison(): ParsedValue {
      let left = parseAdditive();
      while (
        cursor < tokens.length &&
        tokens[cursor]?.type === 'OP' &&
        ['=', '<>', '<', '>', '<=', '>='].includes(tokens[cursor]!.value)
      ) {
        const op = tokens[cursor]!.value;
        cursor++;
        const right = parseAdditive();
        if (op === '=') left = formulaEquals(left, right);
        else if (op === '<>') {
          const equality = formulaEquals(left, right);
          left = isFormulaError(equality) ? equality : equality === false;
        } else if (op === '<') left = orderedComparison(left, right, (n) => n < 0);
        else if (op === '>') left = orderedComparison(left, right, (n) => n > 0);
        else if (op === '<=') left = orderedComparison(left, right, (n) => n <= 0);
        else if (op === '>=') left = orderedComparison(left, right, (n) => n >= 0);
      }
      return left;
    }

    function parseAdditive(): ParsedValue {
      let left = parseMultiplicative();
      while (
        cursor < tokens.length &&
        tokens[cursor]?.type === 'OP' &&
        ['+', '-', '&'].includes(tokens[cursor]!.value)
      ) {
        const op = tokens[cursor]!.value;
        cursor++;
        const right = parseMultiplicative();
        if (op === '&') left = firstError(left, right) ?? String(left ?? '') + String(right ?? '');
        else left = addOrSubtract(left, right, op === '-');
      }
      return left;
    }

    function parseMultiplicative(): ParsedValue {
      let left = parsePower();
      while (
        cursor < tokens.length &&
        tokens[cursor]?.type === 'OP' &&
        ['*', '/'].includes(tokens[cursor]!.value)
      ) {
        const op = tokens[cursor]!.value;
        cursor++;
        const right = parsePower();
        left = arithmetic(left, right, op === '*' ? '*' : '/');
      }
      return left;
    }

    function parsePower(): ParsedValue {
      const left = parseUnary();
      if (
        cursor < tokens.length &&
        tokens[cursor]?.type === 'OP' &&
        tokens[cursor]!.value === '^'
      ) {
        cursor++;
        const right = parsePower(); // '^' is right-associative in Excel
        return arithmetic(left, right, '^');
      }
      return left;
    }

    function parseUnary(): ParsedValue {
      let val: ParsedValue;
      if (
        cursor < tokens.length &&
        tokens[cursor]?.type === 'OP' &&
        ['+', '-'].includes(tokens[cursor]!.value)
      ) {
        const op = tokens[cursor]!.value;
        cursor++;
        const v = parseUnary();
        if (isFormulaError(v)) return v;
        const numeric = toNumeric(v);
        if (isFormulaError(numeric)) return numeric;
        val = op === '-' ? -numeric : numeric;
      } else {
        val = parsePrimary();
      }
      // Postfix percent operator: 50% -> 0.5
      while (
        cursor < tokens.length &&
        tokens[cursor]?.type === 'OP' &&
        tokens[cursor]!.value === '%'
      ) {
        if (isFormulaError(val)) return val;
        const numeric = toNumeric(val);
        if (isFormulaError(numeric)) return numeric;
        val = numeric / 100;
        cursor++;
      }
      return val;
    }

    function parsePrimary(): ParsedValue {
      if (cursor >= tokens.length) {
        malformed = true;
        return '#ERROR!';
      }
      const t = tokens[cursor]!;

      // Number literal
      if (t.type === 'NUMBER') {
        cursor++;
        if (!/^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(t.value)) {
          malformed = true;
          return '#ERROR!';
        }
        const numeric = Number(t.value);
        return Number.isFinite(numeric) ? numeric : '#NUM!';
      }

      // String literal
      if (t.type === 'STRING') {
        cursor++;
        return t.value;
      }

      // Boolean
      if (t.type === 'BOOLEAN') {
        cursor++;
        return t.value === 'TRUE';
      }

      if (t.type === 'ERROR') {
        cursor++;
        return t.value as FormulaErrorCode;
      }

      // Parentheses (expr)
      if (t.type === 'LPAREN') {
        depth += 1;
        if (depth > MAX_EVALUATION_DEPTH) {
          malformed = true;
          depth -= 1;
          return '#ERROR!';
        }
        cursor++;
        const val = parseExpression();
        const closed = cursor < tokens.length && tokens[cursor]?.type === 'RPAREN';
        if (closed) cursor++;
        else malformed = true;
        depth -= 1;
        return closed ? val : '#ERROR!';
      }

      // Cell reference (A1, Sheet1!B2, 'My Sheet'!C3)
      if (t.type === 'CELL') {
        cursor++;
        const match = t.value.match(/^(?:(.+)!)?\$?([A-Za-z]+)\$?(\d+)$/);
        if (match) {
          const sheet = t.sheet ?? match[1] ?? context.activeSheet;
          const col = match[2]!.toUpperCase();
          const row = parseInt(match[3]!, 10);
          if (!sheetExists(sheet) || !validReference(col, row)) return '#REF!';
          return context.getCellValue(sheet, col, row);
        }
        return '#REF!';
      }

      // Range reference (A1:B10)
      if (t.type === 'RANGE') {
        cursor++;
        const match = t.value.match(/^(?:(.+)!)?\$?([A-Za-z]+)\$?(\d+):\$?([A-Za-z]+)\$?(\d+)$/);
        if (match) {
          const sheet = t.sheet ?? match[1] ?? context.activeSheet;
          const startCol = match[2]!.toUpperCase();
          const startRow = parseInt(match[3]!, 10);
          const endCol = match[4]!.toUpperCase();
          const endRow = parseInt(match[5]!, 10);
          if (
            !sheetExists(sheet) ||
            !validReference(startCol, startRow) ||
            !validReference(endCol, endRow)
          )
            return '#REF!';
          // A reversed range such as A5:A1 is a malformed reference, not an empty one.
          if (endRow < startRow || columnToIndex(endCol)! < columnToIndex(startCol)!)
            return '#REF!';
          return context.getRangeValues(sheet, startCol, startRow, endCol, endRow);
        }
        return '#REF!';
      }

      // Function call IDENT(args...)
      if (t.type === 'IDENT') {
        const fnName = t.value;
        cursor++;

        if (cursor < tokens.length && tokens[cursor]?.type === 'LPAREN') {
          depth += 1;
          if (depth > MAX_EVALUATION_DEPTH) {
            malformed = true;
            depth -= 1;
            return '#ERROR!';
          }
          cursor++; // consume '('
          const args: ParsedValue[] = [];

          if (cursor < tokens.length && tokens[cursor]?.type !== 'RPAREN') {
            while (cursor < tokens.length) {
              args.push(parseExpression());
              if (cursor < tokens.length && tokens[cursor]?.type === 'COMMA') {
                cursor++; // consume ','
              } else {
                break;
              }
            }
          }

          const closed = cursor < tokens.length && tokens[cursor]?.type === 'RPAREN';
          if (closed) cursor++; // consume ')'
          depth -= 1;
          if (!closed) {
            malformed = true;
            return '#ERROR!';
          }

          const fn = FORMULA_FUNCTIONS[fnName];
          if (fn) {
            const value = fn(...args);
            return typeof value === 'number' && !Number.isFinite(value) ? '#NUM!' : value;
          }
          // The canonical Excel code, so IFERROR/IFNA can actually catch an unknown name.
          return '#NAME?';
        }

        return '#NAME?';
      }

      if (t.type === 'INVALID') {
        cursor++;
        return '#ERROR!';
      }

      malformed = true;
      cursor++;
      return '#ERROR!';
    }

    const result = parseExpression();

    // Trailing tokens mean the formula was malformed. Silently returning the prefix would
    // turn `=1+1)+DROP(A1)` into `2`, which is worse than reporting the problem.
    if (malformed || cursor < tokens.length) return '#ERROR!';

    return result as FormulaValue;
  } catch {
    return '#ERROR!';
  } finally {
    setFormulaDateSystem(previousDateSystem);
  }
}
