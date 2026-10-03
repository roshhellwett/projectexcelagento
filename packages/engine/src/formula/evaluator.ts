import { FORMULA_FUNCTIONS } from './functions.js';
import type { FormulaContext, FormulaValue } from './types.js';

type ParsedValue = FormulaValue | FormulaValue[][];

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
    | 'COMMA';
  value: string;
  sheet?: string;
}

const CELL_REGEX = /^(?:([A-Za-z0-9_]+)!)?\$?([A-Za-z]+)\$?(\d+)$/;
const RANGE_REGEX = /^(?:([A-Za-z0-9_]+)!)?\$?([A-Za-z]+)\$?(\d+):\$?([A-Za-z]+)\$?(\d+)$/;

export function tokenize(formulaStr: string): Token[] {
  let str = formulaStr.trim();
  if (str.startsWith('=')) str = str.slice(1).trim();

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
      while (j < str.length && str[j] !== quote) {
        if (str[j] === '\\' && j + 1 < str.length) {
          text += str[j + 1];
          j += 2;
        } else {
          text += str[j];
          j++;
        }
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
        }
        i = k;
        continue;
      }
      i = j + 1; // closing quote
      tokens.push({ type: 'STRING', value: text });
      continue;
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
      // Unrecognized single character - skip and advance cursor to prevent infinite loop
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

function formulaEquals(a: ParsedValue, b: ParsedValue): boolean {
  if (a === null && b === null) return true;
  if (a === null || b === null) return b === '' || a === '';
  if (typeof a === 'number' && typeof b === 'number') return a === b;
  if (typeof a === 'boolean' || typeof b === 'boolean') return a === b;
  const an = typeof a === 'string' && a.trim() !== '' ? Number(a) : NaN;
  const bn = typeof b === 'string' && b.trim() !== '' ? Number(b) : NaN;
  if (!isNaN(an) && !isNaN(bn)) return an === bn;
  if (typeof a === 'number' && !isNaN(bn as number)) return a === bn;
  if (typeof b === 'number' && !isNaN(an as number)) return b === an;
  return String(a) === String(b);
}

function compareOrdered(a: ParsedValue, b: ParsedValue): number {
  const rank = (v: ParsedValue): number =>
    typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 3;
  const aNum =
    typeof a === 'number'
      ? a
      : typeof a === 'string' && a.trim() !== '' && !isNaN(Number(a))
        ? Number(a)
        : NaN;
  const bNum =
    typeof b === 'number'
      ? b
      : typeof b === 'string' && b.trim() !== '' && !isNaN(Number(b))
        ? Number(b)
        : NaN;
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
  try {
    const tokens = tokenize(formula);
    if (tokens.length === 0) return null;

    let cursor = 0;

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
        else if (op === '<>') left = !formulaEquals(left, right);
        else if (op === '<') left = compareOrdered(left, right) < 0;
        else if (op === '>') left = compareOrdered(left, right) > 0;
        else if (op === '<=') left = compareOrdered(left, right) <= 0;
        else if (op === '>=') left = compareOrdered(left, right) >= 0;
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
        if (op === '+') left = Number(left) + Number(right);
        else if (op === '-') left = Number(left) - Number(right);
        else if (op === '&') left = String(left ?? '') + String(right ?? '');
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
        if (op === '*') left = Number(left) * Number(right);
        else if (op === '/') left = Number(right) === 0 ? '#DIV/0!' : Number(left) / Number(right);
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
        return Math.pow(Number(left), Number(right));
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
        val = op === '-' ? -Number(v) : Number(v);
      } else {
        val = parsePrimary();
      }
      // Postfix percent operator: 50% -> 0.5
      while (
        cursor < tokens.length &&
        tokens[cursor]?.type === 'OP' &&
        tokens[cursor]!.value === '%'
      ) {
        val = Number(val) / 100;
        cursor++;
      }
      return val;
    }

    function parsePrimary(): ParsedValue {
      if (cursor >= tokens.length) return null;
      const t = tokens[cursor]!;

      // Number literal
      if (t.type === 'NUMBER') {
        cursor++;
        return parseFloat(t.value);
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

      // Parentheses (expr)
      if (t.type === 'LPAREN') {
        cursor++;
        const val = parseExpression();
        if (cursor < tokens.length && tokens[cursor]?.type === 'RPAREN') {
          cursor++;
        }
        return val;
      }

      // Cell reference (A1, Sheet1!B2, 'My Sheet'!C3)
      if (t.type === 'CELL') {
        cursor++;
        const match = t.value.match(/^(?:(.+)!)?\$?([A-Za-z]+)\$?(\d+)$/);
        if (match) {
          const sheet = t.sheet ?? match[1] ?? context.activeSheet;
          const col = match[2]!.toUpperCase();
          const row = parseInt(match[3]!, 10);
          return context.getCellValue(sheet, col, row);
        }
        return null;
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
          return context.getRangeValues(sheet, startCol, startRow, endCol, endRow);
        }
        return [];
      }

      // Function call IDENT(args...)
      if (t.type === 'IDENT') {
        const fnName = t.value;
        cursor++;

        if (cursor < tokens.length && tokens[cursor]?.type === 'LPAREN') {
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

          if (cursor < tokens.length && tokens[cursor]?.type === 'RPAREN') {
            cursor++; // consume ')'
          }

          const fn = FORMULA_FUNCTIONS[fnName];
          if (fn) {
            return fn(...args);
          }
          return `#NAME? (${fnName})`;
        }

        return `#NAME? (${fnName})`;
      }

      cursor++;
      return null;
    }

    return parseExpression() as FormulaValue;
  } catch {
    return '#ERROR!';
  }
}
