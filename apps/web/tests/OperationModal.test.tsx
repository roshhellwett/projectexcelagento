// @vitest-environment jsdom
import { describe, expect, it, vi } from 'vitest';
import { fireEvent, render, screen } from '@testing-library/react';
import { createCell } from '@excel-agent/engine';

import { OperationModal } from '../src/components/OperationModal.js';

describe('OperationModal', () => {
  it('formats the active selection without requiring JSON arguments', () => {
    const onExecute = vi.fn();
    render(
      <OperationModal
        isOpen
        onClose={vi.fn()}
        workbook={{
          sheets: [{ name: 'Sheet1', rows: [[createCell('Name'), createCell('Amount')]] }],
        }}
        activeSheetName="Sheet1"
        selection={{ startRow: 1, endRow: 1, startColIdx: 0, endColIdx: 1 }}
        onExecute={onExecute}
      />,
    );

    fireEvent.change(screen.getByLabelText('Choose a tool'), {
      target: { value: 'format_cells' },
    });
    fireEvent.click(screen.getByRole('button', { name: 'Bold' }));
    fireEvent.change(screen.getByLabelText('Number format'), {
      target: { value: '$#,##0.00' },
    });
    fireEvent.click(screen.getByRole('button', { name: 'Apply to selection' }));

    expect(onExecute).toHaveBeenCalledWith('format_cells', {
      sheet: 'Sheet1',
      startRow: 1,
      endRow: 1,
      startColumn: 'A',
      endColumn: 'B',
      style: { bold: true },
      numberFormat: '$#,##0.00',
    });
  });

  it('applies a filter to the view without mutating workbook rows', () => {
    const onExecute = vi.fn();
    const onApplyFilter = vi.fn();
    render(
      <OperationModal
        isOpen
        onClose={vi.fn()}
        workbook={{ sheets: [{ name: 'Sheet1', rows: [[createCell('Status')]] }] }}
        activeSheetName="Sheet1"
        onExecute={onExecute}
        onApplyFilter={onApplyFilter}
      />,
    );

    fireEvent.change(screen.getByLabelText('Choose a tool'), {
      target: { value: 'filter_rows' },
    });
    fireEvent.change(screen.getByLabelText('Comparison Value'), {
      target: { value: 'Completed' },
    });
    fireEvent.click(screen.getByRole('button', { name: 'Apply to selection' }));

    expect(onApplyFilter).toHaveBeenCalledWith({
      sheet: 'Sheet1',
      column: 'A',
      operator: 'equals',
      value: 'Completed',
    });
    expect(onExecute).not.toHaveBeenCalled();
  });

  it('exposes the complete registered operation catalog through validated JSON arguments', () => {
    const onExecute = vi.fn();
    render(
      <OperationModal
        isOpen
        onClose={vi.fn()}
        workbook={{ sheets: [{ name: 'Sheet1', rows: [[createCell('Name')]] }] }}
        activeSheetName="Sheet1"
        operationCatalog={[
          {
            name: 'create_sheet',
            description: 'Create a worksheet.',
            example: { sheetName: 'Summary' },
          },
        ]}
        onExecute={onExecute}
      />,
    );

    fireEvent.change(screen.getByLabelText('Choose a tool'), {
      target: { value: 'create_sheet' },
    });
    fireEvent.change(screen.getByLabelText('Operation arguments (JSON)'), {
      target: { value: '{"sheetName":"Summary"}' },
    });
    fireEvent.click(screen.getByRole('button', { name: 'Apply to selection' }));

    expect(onExecute).toHaveBeenCalledWith('create_sheet', { sheetName: 'Summary' });
  });
});
