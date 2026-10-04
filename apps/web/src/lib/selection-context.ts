import type { Sheet } from '@excel-agent/engine';
import { indexToColumn } from '@excel-agent/engine';

export type SelectionKind = 'cell' | 'row' | 'column';

export interface CellSelection {
  kind: SelectionKind;
  sheetName: string;
  /** e.g. "C5", "Row 5", "Column C" */
  label: string;
  /** Instruction prefix injected into the agent query. */
  summary: string;
}

/**
 * Renders a cell value for inclusion in the agent query.
 *
 * This summary is concatenated into the query string that the deterministic keyword planner
 * scans, so a cell whose text happens to be "delete all rows" would otherwise steer the
 * planner by itself. Collapsing whitespace and stripping the separators that structure a
 * request keeps a cell's words from being read as a command, while the value stays visible to
 * the user and to the model.
 */
const preview = (value: unknown): string => {
  if (value === null || value === undefined || value === '') return '(empty)';
  const text = value instanceof Date ? value.toISOString().slice(0, 10) : String(value);
  const inert = text
    .replace(/[\r\n]+/g, ' ')
    .replace(/[;:`"'\\]/g, ' ')
    .trim();
  return inert.length > 40 ? `${inert.slice(0, 40)}...` : inert;
};

export function describeCellSelection(
  kind: SelectionKind,
  sheet: Sheet,
  rowIndex: number, // 0-based row index into sheet.rows
  colIdx: number, // 0-based column index
): CellSelection {
  const colLetter = indexToColumn(colIdx);
  const rowNumber = rowIndex + 1;
  const cell = sheet.rows[rowIndex]?.[colIdx];
  if (kind === 'cell') {
    const address = `${colLetter}${rowNumber}`;
    return {
      kind,
      sheetName: sheet.name,
      label: address,
      summary: `Selected cell ${sheet.name}!${address} (current value: ${preview(cell?.value)}${cell?.formula ? `, formula: ${cell.formula}` : ''}).`,
    };
  }
  if (kind === 'row') {
    const row = sheet.rows[rowIndex] ?? [];
    const values = row
      .slice(0, 6)
      .map((c, idx) => `${indexToColumn(idx)}${rowNumber}=${preview(c.value)}`)
      .join(', ');
    return {
      kind,
      sheetName: sheet.name,
      label: `Row ${rowNumber}`,
      summary: `Selected row ${rowNumber} on sheet "${sheet.name}" (${values}${row.length > 6 ? ', ...' : ''}).`,
    };
  }
  const values = sheet.rows
    .slice(0, 6)
    .map((row, ridx) => `${colLetter}${ridx + 1}=${preview(row[colIdx]?.value)}`)
    .join(', ');
  return {
    kind,
    sheetName: sheet.name,
    label: `Column ${colLetter}`,
    summary: `Selected column ${colLetter} on sheet "${sheet.name}" (${values}${sheet.rows.length > 6 ? ', ...' : ''}).`,
  };
}
