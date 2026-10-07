// @vitest-environment jsdom
import { describe, expect, it, vi } from 'vitest';
import { fireEvent, render, screen } from '@testing-library/react';
import { createCell } from '@excel-agent/engine';

import { OperationModal } from '../src/components/OperationModal.js';

describe('OperationModal', () => {
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

    fireEvent.change(screen.getByLabelText('Select Operation'), {
      target: { value: 'create_sheet' },
    });
    fireEvent.change(screen.getByLabelText('Operation arguments (JSON)'), {
      target: { value: '{"sheetName":"Summary"}' },
    });
    fireEvent.click(screen.getByRole('button', { name: 'Execute Operation' }));

    expect(onExecute).toHaveBeenCalledWith('create_sheet', { sheetName: 'Summary' });
  });
});
