// @vitest-environment jsdom
import { describe, expect, it, vi } from 'vitest';
import { fireEvent, render, screen } from '@testing-library/react';

import { WorkspaceRibbon } from '../src/components/WorkspaceRibbon.js';

function renderRibbon() {
  return render(
    <WorkspaceRibbon
      activeSheetName="Orders"
      isProcessing={false}
      onManualAction={vi.fn()}
      onOpenOperationModal={vi.fn()}
      onUploadFile={vi.fn()}
      onExport={vi.fn()}
      onUndo={vi.fn()}
      onRedo={vi.fn()}
      canUndo
      canRedo={false}
    />,
  );
}

describe('WorkspaceRibbon', () => {
  it('keeps the requested Excel workspace tabs discoverable in one surface', () => {
    renderRibbon();

    for (const tab of [
      'Home',
      'Insert',
      'Draw',
      'Page Layout',
      'Formulas',
      'Data',
      'Review',
      'Automate',
      'Help',
    ]) {
      expect(screen.getByRole('tab', { name: tab })).toBeInTheDocument();
    }

    expect(screen.getByRole('button', { name: /Currency/i })).toBeInTheDocument();
    expect(screen.getByRole('button', { name: /AutoSum/i })).toBeInTheDocument();
    expect(screen.getByRole('button', { name: /Sort & Filter/i })).toBeInTheDocument();
  });

  it('routes an active ribbon action through the manual action path', () => {
    const onManualAction = vi.fn();
    render(
      <WorkspaceRibbon
        activeSheetName="Orders"
        isProcessing={false}
        onManualAction={onManualAction}
        onOpenOperationModal={vi.fn()}
        onUploadFile={vi.fn()}
        onExport={vi.fn()}
        onUndo={vi.fn()}
        onRedo={vi.fn()}
        canUndo={false}
        canRedo={false}
      />,
    );

    fireEvent.click(screen.getByRole('button', { name: /Currency/i }));

    expect(onManualAction).toHaveBeenCalledWith('currency');
  });

  it('routes the Borders command through the local formatting action path', () => {
    const onManualAction = vi.fn();
    render(
      <WorkspaceRibbon
        activeSheetName="Orders"
        isProcessing={false}
        onManualAction={onManualAction}
        onOpenOperationModal={vi.fn()}
        onUploadFile={vi.fn()}
        onExport={vi.fn()}
        onUndo={vi.fn()}
        onRedo={vi.fn()}
        canUndo={false}
        canRedo={false}
      />,
    );

    fireEvent.click(screen.getByRole('button', { name: 'Borders' }));

    expect(onManualAction).toHaveBeenCalledWith('borders');
  });

  it('shows unsupported workbook-fidelity controls without pretending they mutate files', () => {
    renderRibbon();
    fireEvent.click(screen.getByRole('tab', { name: 'Insert' }));

    const pictures = screen.getByRole('button', { name: 'Pictures' });
    expect(pictures).toBeDisabled();
    expect(pictures).toHaveAttribute('title', expect.stringContaining('drawing layer'));
  });

  it('keeps search and help clicks on the manual operation surface', () => {
    const onManualAction = vi.fn();
    render(
      <WorkspaceRibbon
        activeSheetName="Orders"
        isProcessing={false}
        onManualAction={onManualAction}
        onOpenOperationModal={vi.fn()}
        onUploadFile={vi.fn()}
        onExport={vi.fn()}
        onUndo={vi.fn()}
        onRedo={vi.fn()}
        canUndo={false}
        canRedo={false}
      />,
    );

    fireEvent.click(screen.getByRole('button', { name: /Find & Select/i }));
    fireEvent.click(screen.getByRole('tab', { name: 'Help' }));
    fireEvent.click(screen.getByRole('button', { name: /Search Help/i }));

    expect(onManualAction).toHaveBeenNthCalledWith(1, 'find-select');
    expect(onManualAction).toHaveBeenNthCalledWith(2, 'search-help');
  });

  it('triggers font, size, and insert actions from the Home tab', () => {
    const onManualAction = vi.fn();
    render(
      <WorkspaceRibbon
        activeSheetName="Orders"
        isProcessing={false}
        onManualAction={onManualAction}
        onOpenOperationModal={vi.fn()}
        onUploadFile={vi.fn()}
        onExport={vi.fn()}
        onUndo={vi.fn()}
        onRedo={vi.fn()}
        canUndo={false}
        canRedo={false}
      />,
    );

    fireEvent.click(screen.getByRole('button', { name: 'Font' }));
    expect(onManualAction).toHaveBeenCalledWith('font');

    fireEvent.click(screen.getByRole('button', { name: 'Size' }));
    expect(onManualAction).toHaveBeenCalledWith('font-size');

    fireEvent.click(screen.getByRole('button', { name: 'Insert' }));
    expect(onManualAction).toHaveBeenCalledWith('insert');
  });

  it('triggers interactive tools on Page Layout, Formulas, and Draw tabs', () => {
    const onManualAction = vi.fn();
    render(
      <WorkspaceRibbon
        activeSheetName="Orders"
        isProcessing={false}
        onManualAction={onManualAction}
        onOpenOperationModal={vi.fn()}
        onUploadFile={vi.fn()}
        onExport={vi.fn()}
        onUndo={vi.fn()}
        onRedo={vi.fn()}
        canUndo={false}
        canRedo={false}
      />,
    );

    // Draw tab
    fireEvent.click(screen.getByRole('tab', { name: 'Draw' }));
    fireEvent.click(screen.getByRole('button', { name: 'Pen' }));
    expect(onManualAction).toHaveBeenCalledWith('pen-black');

    // Page Layout tab
    fireEvent.click(screen.getByRole('tab', { name: 'Page Layout' }));
    fireEvent.click(screen.getByRole('button', { name: 'Themes' }));
    expect(onManualAction).toHaveBeenCalledWith('themes');
    fireEvent.click(screen.getByRole('button', { name: 'Gridlines' }));
    expect(onManualAction).toHaveBeenCalledWith('gridlines');

    // Formulas tab
    fireEvent.click(screen.getByRole('tab', { name: 'Formulas' }));
    fireEvent.click(screen.getByRole('button', { name: 'Show Formulas' }));
    expect(onManualAction).toHaveBeenCalledWith('show-formulas');
  });
});
