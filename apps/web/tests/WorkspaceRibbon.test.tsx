// @vitest-environment jsdom
import { describe, expect, it, vi } from 'vitest';
import { fireEvent, render, screen } from '@testing-library/react';

import { WorkspaceRibbon } from '../src/components/WorkspaceRibbon.js';

function renderRibbon() {
  return render(
    <WorkspaceRibbon
      activeSheetName="Orders"
      isProcessing={false}
      onRunPrompt={vi.fn()}
      onOpenOperationModal={vi.fn()}
      onOpenCommandPalette={vi.fn()}
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

  it('routes an active ribbon action through the agent request path', () => {
    const onRunPrompt = vi.fn();
    render(
      <WorkspaceRibbon
        activeSheetName="Orders"
        isProcessing={false}
        onRunPrompt={onRunPrompt}
        onOpenOperationModal={vi.fn()}
        onOpenCommandPalette={vi.fn()}
        onUploadFile={vi.fn()}
        onExport={vi.fn()}
        onUndo={vi.fn()}
        onRedo={vi.fn()}
        canUndo={false}
        canRedo={false}
      />,
    );

    fireEvent.click(screen.getByRole('button', { name: /Currency/i }));

    expect(onRunPrompt).toHaveBeenCalledWith(expect.stringContaining('currency'));
    expect(onRunPrompt).toHaveBeenCalledWith(expect.stringContaining('Orders'));
  });

  it('shows unsupported workbook-fidelity controls without pretending they mutate files', () => {
    renderRibbon();
    fireEvent.click(screen.getByRole('tab', { name: 'Insert' }));

    const pictures = screen.getByRole('button', { name: 'Pictures' });
    expect(pictures).toBeDisabled();
    expect(pictures).toHaveAttribute('title', expect.stringContaining('drawing layer'));
  });
});
