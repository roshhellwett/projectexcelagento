import { z } from 'zod';

import { runInvariants } from './invariants.js';
import {
  issue,
  previewForTransition,
  transitionResult,
  validResult,
  validateSheet,
} from './operation-utils.js';
import type {
  CellRange,
  CellStyle,
  Operation,
  OperationResult,
  Preview,
  Report,
  ValidationResult,
  Workbook,
} from './types.js';
import { cloneCell, cloneWorkbook, columnToIndex, getSheet, maxColumnCount } from './workbook.js';

const column = z
  .string()
  .trim()
  .regex(/^[A-Za-z]+$/);

export const formatCellsArgsSchema = z.object({
  sheet: z.string().trim().min(1),
  startRow: z.number().int().positive(),
  endRow: z.number().int().positive(),
  startColumn: column,
  endColumn: column,
  style: z
    .object({
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      underline: z.boolean().optional(),
      fillColor: z
        .string()
        .trim()
        .regex(/^#[0-9a-f]{6}$/i)
        .optional(),
      fontColor: z
        .string()
        .trim()
        .regex(/^#[0-9a-f]{6}$/i)
        .optional(),
      horizontalAlignment: z.enum(['left', 'center', 'right']).optional(),
      verticalAlignment: z.enum(['top', 'middle', 'bottom']).optional(),
      wrapText: z.boolean().optional(),
    })
    .partial(),
  numberFormat: z.string().optional(),
});
export type FormatCellsArgs = z.infer<typeof formatCellsArgsSchema>;

function targetRange(_workbook: Workbook, args: FormatCellsArgs): CellRange {
  return {
    sheet: args.sheet,
    startRow: args.startRow,
    endRow: args.endRow,
    startColumn: args.startColumn.toUpperCase(),
    endColumn: args.endColumn.toUpperCase(),
  };
}

function validateFormatCells(workbook: Workbook, args: FormatCellsArgs): ValidationResult {
  const errors = validateSheet(workbook, args.sheet);
  const sheet = getSheet(workbook, args.sheet);
  const startColumn = columnToIndex(args.startColumn);
  const endColumn = columnToIndex(args.endColumn);
  if (startColumn === undefined || endColumn === undefined || endColumn < startColumn) {
    errors.push(issue('invalid-range', 'The formatting range must have valid ordered columns.'));
  }
  if (args.endRow < args.startRow) {
    errors.push(issue('invalid-range', 'The formatting range must have ordered rows.'));
  }
  if (sheet && args.endRow > sheet.rows.length) {
    errors.push(issue('invalid-range', 'The formatting range extends beyond the used rows.'));
  }
  if (endColumn !== undefined && endColumn >= maxColumnCount(sheet?.rows ?? [])) {
    errors.push(issue('invalid-range', 'The formatting range extends beyond the used columns.'));
  }
  const area =
    startColumn === undefined || endColumn === undefined
      ? Infinity
      : (args.endRow - args.startRow + 1) * (endColumn - startColumn + 1);
  if (area > 20_000) {
    errors.push(
      issue('range-too-large', 'A single formatting gesture may affect at most 20,000 cells.'),
    );
  }
  if (Object.keys(args.style).length === 0 && args.numberFormat === undefined) {
    errors.push(issue('empty-format', 'Choose at least one formatting property.'));
  }
  return errors.length === 0 ? validResult() : { valid: false, errors, warnings: [] };
}

function applyFormatCells(workbook: Workbook, args: FormatCellsArgs): OperationResult {
  const before = cloneWorkbook(workbook);
  const after = cloneWorkbook(workbook);
  const sheet = getSheet(after, args.sheet);
  const startColumn = columnToIndex(args.startColumn) ?? 0;
  const endColumn = columnToIndex(args.endColumn) ?? startColumn;
  const style = args.style as CellStyle;
  for (let row = args.startRow - 1; row < args.endRow; row += 1) {
    for (let col = startColumn; col <= endColumn; col += 1) {
      const cell = sheet?.rows[row]?.[col];
      if (!cell) continue;
      const next = cloneCell(cell);
      next.style = { ...(next.style ?? {}), ...style };
      if (args.numberFormat !== undefined) next.numberFormat = args.numberFormat;
      if (Object.keys(next.style).length === 0) delete next.style;
      if (sheet?.rows[row]) sheet.rows[row]![col] = next;
    }
  }
  const range = targetRange(workbook, args);
  const preview = previewForTransition(before, after, [range]);
  const report: Report = {
    affectedCells: preview.affectedCells,
    skippedCells: 0,
    unchangedCells: Math.max(
      0,
      before.sheets.reduce((total, item) => total + item.rows.flat().length, 0) -
        preview.affectedCells,
    ),
    warnings: [],
  };
  return transitionResult(before, after, report);
}

function invalidPreview(
  workbook: Workbook,
  args: FormatCellsArgs,
  errors: ValidationResult['errors'],
): Preview {
  return previewForTransition(workbook, workbook, [targetRange(workbook, args)], [], errors);
}

export const formatCellsOperation: Operation<FormatCellsArgs> = {
  name: 'format_cells',
  schema: formatCellsArgsSchema,
  targetRanges: (workbook, args) => [targetRange(workbook, args)],
  validate: validateFormatCells,
  preview(workbook, args) {
    const validation = validateFormatCells(workbook, args);
    return validation.valid
      ? previewForTransition(workbook, applyFormatCells(workbook, args).workbook, [
          targetRange(workbook, args),
        ])
      : invalidPreview(workbook, args, validation.errors);
  },
  apply: applyFormatCells,
  invariants(before, after, args) {
    return runInvariants(before, after, {
      targetRanges: [targetRange(before, args)],
      rowCountUnchanged: true,
    });
  },
};
