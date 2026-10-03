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

const preview = (value: unknown): string => {
  if (value === null || value === undefined || value === '') return '(empty)';
  const s = String(value);
  return s.length > 40 ? `${s.slice(0, 40)}...` : s;
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
