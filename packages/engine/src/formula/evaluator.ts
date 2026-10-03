import { FORMULA_FUNCTIONS } from './functions.js';
import type { FormulaContext, FormulaValue } from './types.js';

interface Token {
  type: 'NUMBER' | 'STRING' | 'BOOLEAN' | 'IDENT' | 'CELL' | 'RANGE' | 'OP' | 'LPAREN' | 'RPAREN' | 'COMMA';
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

    // String literals
    if (ch === '"' || ch === "'") {
      const quote = ch;
      let text = '';
      i++;
      while (i < str.length && str[i] !== quote) {
        if (str[i] === '\\' && i + 1 < str.length) {
          text += str[i + 1];
          i += 2;
        } else {
          text += str[i];
          i++;
        }
      }
      i++; // closing quote
      tokens.push({ type: 'STRING', value: text });
      continue;
    }

    // Numbers
    if (/[0-9]/.test(ch) || (ch === '.' && /[0-9]/.test(str[i + 1] ?? ''))) {
      let numStr = '';
      while (i < str.length && /[0-9.]/.test(str[i]!)) {
        numStr += str[i];
        i++;
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

export function evaluateFormula(formula: string, context: FormulaContext): FormulaValue {
  try {
    const tokens = tokenize(formula);
    if (tokens.length === 0) return null;

    let cursor = 0;

    function parseExpression(): any {
      return parseLogicalOr();
    }

    function parseLogicalOr(): any {
      let left = parseComparison();
      // In Excel, OR is usually a function, but handle infix comparison / logical
      return left;
    }

    function parseComparison(): any {
      let left = parseAdditive();
      while (cursor < tokens.length && tokens[cursor]?.type === 'OP' && ['=', '<>', '<', '>', '<=', '>='].includes(tokens[cursor]!.value)) {
        const op = tokens[cursor]!.value;
        cursor++;
        const right = parseAdditive();
        if (op === '=') left = left === right;
        else if (op === '<>') left = left !== right;
        else if (op === '<') left = left < right;
        else if (op === '>') left = left > right;
        else if (op === '<=') left = left <= right;
        else if (op === '>=') left = left >= right;
      }
      return left;
    }

    function parseAdditive(): any {
      let left = parseMultiplicative();
      while (cursor < tokens.length && tokens[cursor]?.type === 'OP' && ['+', '-', '&'].includes(tokens[cursor]!.value)) {
        const op = tokens[cursor]!.value;
        cursor++;
        const right = parseMultiplicative();
        if (op === '+') left = Number(left) + Number(right);
        else if (op === '-') left = Number(left) - Number(right);
        else if (op === '&') left = String(left ?? '') + String(right ?? '');
      }
      return left;
    }

    function parseMultiplicative(): any {
      let left = parsePower();
      while (cursor < tokens.length && tokens[cursor]?.type === 'OP' && ['*', '/', '%'].includes(tokens[cursor]!.value)) {
        const op = tokens[cursor]!.value;
        cursor++;
        const right = parsePower();
        if (op === '*') left = Number(left) * Number(right);
        else if (op === '/') left = Number(right) === 0 ? '#DIV/0!' : Number(left) / Number(right);
        else if (op === '%') left = Number(left) % Number(right);
      }
      return left;
    }

    function parsePower(): any {
      let left = parseUnary();
      while (cursor < tokens.length && tokens[cursor]?.type === 'OP' && tokens[cursor]!.value === '^') {
        cursor++;
        const right = parseUnary();
        left = Math.pow(Number(left), Number(right));
      }
      return left;
    }

    function parseUnary(): any {
      if (cursor < tokens.length && tokens[cursor]?.type === 'OP' && ['+', '-'].includes(tokens[cursor]!.value)) {
        const op = tokens[cursor]!.value;
        cursor++;
        const val = parseUnary();
        return op === '-' ? -Number(val) : Number(val);
      }
      return parsePrimary();
    }

    function parsePrimary(): any {
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

      // Cell reference (A1, Sheet1!B2)
      if (t.type === 'CELL') {
        cursor++;
        const match = t.value.match(CELL_REGEX);
        if (match) {
          const sheet = match[1] ?? context.activeSheet;
          const col = match[2]!.toUpperCase();
          const row = parseInt(match[3]!, 10);
          return context.getCellValue(sheet, col, row);
        }
        return null;
      }

      // Range reference (A1:B10)
      if (t.type === 'RANGE') {
        cursor++;
        const match = t.value.match(RANGE_REGEX);
        if (match) {
          const sheet = match[1] ?? context.activeSheet;
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
          const args: any[] = [];

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

    return parseExpression();
  } catch (err) {
    return '#ERROR!';
  }
}
