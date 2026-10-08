import { describe, expect, it } from 'vitest';
import {
  applyOperation,
  applyOperationPlan,
  createCell,
  HistoryStack,
  MAX_OPERATION_PLAN_STEPS,
  type Workbook,
} from '../src/index.js';

const workbook = (): Workbook => ({
  sheets: [
    { name: 'Data', rows: [[createCell('Name')], [createCell(' Ada ')], [createCell(' Ada ')]] },
  ],
});
const trim = { operation: 'normalize_text', args: { sheet: 'Data', columns: ['A'], trim: true } };
const dedupe = {
  operation: 'delete_duplicates',
  args: { sheet: 'Data', columns: ['A'], headerRow: 1 },
};

describe('atomic operation plans', () => {
  it('requires confirmation without changing workbook or history', () => {
    const before = workbook();
    const history = new HistoryStack(before);
    const result = applyOperationPlan(before, [trim, dedupe], { history });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.code).toBe('confirmation-required');
    expect(result.workbook).toEqual(before);
    expect(history.length).toBe(0);
    expect(history.canRedo).toBe(false);
  });

  it('commits all steps as one undoable change', () => {
    const before = workbook();
    const history = new HistoryStack(before);
    const result = applyOperationPlan(before, [trim, dedupe], { history, confirmed: true });
    expect(result.ok).toBe(true);
    expect(result.workbook.sheets[0]?.rows).toHaveLength(2);
    expect(result.workbook.sheets[0]?.rows[1]?.[0]?.value).toBe('Ada');
    expect(before).toEqual(workbook());
    expect(history.length).toBe(1);
    expect(history.undo()).toEqual(before);
    expect(history.redo()).toEqual(result.workbook);
  });

  it('preserves an existing redo branch and history budget when a later step fails', () => {
    const before = workbook();
    const history = new HistoryStack(before, { maxEntries: 1, maxRetainedCells: 1 });
    const edited = applyOperation(
      before,
      'normalize_text',
      { ...trim.args, case: 'upper' },
      { history },
    );
    expect(edited.ok).toBe(true);
    history.undo();
    const historyBefore = history.history;
    const result = applyOperationPlan(
      before,
      [trim, { operation: 'delete_column', args: { sheet: 'Data', column: 'ZZ' } }],
      { history, confirmed: true },
    );
    expect(result.ok).toBe(false);
    expect(result.workbook).toEqual(before);
    expect(history.history).toEqual(historyBefore);
    expect(history.position).toBe(0);
    expect(history.redo()).toEqual(edited.workbook);
  });

  it('validates later steps against sheets created by earlier steps', () => {
    const result = applyOperationPlan(workbook(), [
      { operation: 'duplicate_sheet', args: { sheet: 'Data', targetSheet: 'Copy' } },
      { ...trim, args: { ...trim.args, sheet: 'Copy' } },
    ]);
    expect(result.ok).toBe(true);
    expect(result.workbook.sheets[1]?.rows[1]?.[0]?.value).toBe('Ada');
  });

  it('rejects an oversized plan before simulating any operation', () => {
    const before = workbook();
    const result = applyOperationPlan(
      before,
      Array.from({ length: MAX_OPERATION_PLAN_STEPS + 1 }, () => trim),
      { confirmed: true },
    );
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.messages[0]).toContain('no more than');
    expect(before).toEqual(workbook());
  });
});
