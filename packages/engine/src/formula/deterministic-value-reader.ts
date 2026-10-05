import type { Workbook } from '../types.js';
import { columnToIndex } from '../workbook.js';
import { FORMULA_FUNCTIONS } from './functions.js';
import { isFormulaError } from './errors.js';
import { tokenize } from './evaluator.js';
import type { FormulaValue } from './types.js';
import { createWorkbookValueReader } from './workbook-reader.js';

export type DeterministicValueReason = 'unsupported' | 'volatile' | 'error' | null;
export interface DeterministicValue { value: FormulaValue; reason: DeterministicValueReason }
export const DETERMINISTIC_READER_LIMITS = Object.freeze({ dependencyEdges: 200000, rangeFormulaScans: 5000000 });

/**
 * Read live values for repeatable analyst jobs, never cached formula results. Unsupported,
 * clock/random formulas and their known dependents are excluded even behind IFERROR.
 * Dependency analysis is bounded: if its budget is exceeded all formula cells are excluded,
 * rather than partially analysing the graph and accidentally evaluating an unverified formula.
 */
export function createDeterministicWorkbookValueReader(workbook: Workbook): (sheet: string, column: number, row: number) => DeterministicValue {
  const key = (sheet: string, column: number, row: number) => JSON.stringify([sheet.toLowerCase(), column, row]);
  const formulas = workbook.sheets.flatMap((sheet) => sheet.rows.flatMap((row, rowIndex) => row.flatMap((cell, column) =>
    cell?.formula !== undefined || cell?.type === 'formula' ? [{ sheet: sheet.name, column, row: rowIndex + 1, cell, key: key(sheet.name, column, rowIndex + 1) }] : [])));
  const formulaKeys = new Set(formulas.map((formula) => formula.key));
  const bySheet = new Map<string, typeof formulas>();
  for (const formula of formulas) { const group = bySheet.get(formula.sheet.toLowerCase()) ?? []; group.push(formula); bySheet.set(formula.sheet.toLowerCase(), group); }
  const blocked = new Map<string, Exclude<DeterministicValueReason, 'error' | null>>();
  const dependents = new Map<string, Set<string>>();
  const volatile = new Set(['TODAY', 'NOW', 'RAND', 'RANDBETWEEN']);
  let edges = 0; let scans = 0; let exhausted = false;
  analysis: for (const formula of formulas) {
    if (!formula.cell.formula?.trim()) blocked.set(formula.key, 'unsupported');
    const tokens = tokenize(formula.cell.formula ?? '');
    for (const [index, token] of tokens.entries()) {
      if (token.type === 'IDENT' && tokens[index + 1]?.type === 'LPAREN') {
        if (volatile.has(token.value)) blocked.set(formula.key, 'volatile');
        else if (!Object.hasOwn(FORMULA_FUNCTIONS, token.value) && !blocked.has(formula.key)) blocked.set(formula.key, 'unsupported');
      }
      if (token.type !== 'CELL' && token.type !== 'RANGE') continue;
      const sheetName = token.sheet ?? formula.sheet;
      const address = token.value.slice(token.value.lastIndexOf('!') + 1).replace(/\$/g, '');
      const match = /^([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$/.exec(address);
      if (!match) continue;
      const fromColumn = columnToIndex(match[1]!)!; const fromRow = Number(match[2]);
      const toColumn = match[3] ? columnToIndex(match[3])! : fromColumn; const toRow = match[4] ? Number(match[4]) : fromRow;
      const dependencies: string[] = [];
      if (token.type === 'CELL') { const dependency = key(sheetName, fromColumn, fromRow); if (formulaKeys.has(dependency)) dependencies.push(dependency); }
      else for (const other of bySheet.get(sheetName.toLowerCase()) ?? []) {
        scans += 1;
        if (scans > DETERMINISTIC_READER_LIMITS.rangeFormulaScans) { exhausted = true; break analysis; }
        if (other.column >= Math.min(fromColumn, toColumn) && other.column <= Math.max(fromColumn, toColumn) && other.row >= Math.min(fromRow, toRow) && other.row <= Math.max(fromRow, toRow)) dependencies.push(other.key);
      }
      for (const dependency of dependencies) {
        edges += 1;
        if (edges > DETERMINISTIC_READER_LIMITS.dependencyEdges) { exhausted = true; break analysis; }
        const entries = dependents.get(dependency) ?? new Set<string>(); entries.add(formula.key); dependents.set(dependency, entries);
      }
    }
  }
  if (exhausted) for (const formula of formulas) blocked.set(formula.key, 'unsupported');
  const queue = [...blocked.keys()];
  for (let index = 0; index < queue.length; index += 1) {
    const source = queue[index]!;
    for (const dependent of dependents.get(source) ?? []) { if (blocked.has(dependent)) continue; blocked.set(dependent, blocked.get(source)!); queue.push(dependent); }
  }
  const read = createWorkbookValueReader(workbook);
  return (sheet, column, row) => {
    const reason = blocked.get(key(sheet, column, row));
    if (reason) return { value: '#ERROR!', reason };
    try { const value = read(sheet, column, row); return { value, reason: isFormulaError(value) ? 'error' : null }; }
    catch { return { value: '#ERROR!', reason: 'error' }; }
  };
}
