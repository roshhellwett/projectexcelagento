// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { render, screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import type { HistoryEntry } from '@excel-agent/engine';

import { HistoryDrawer } from '../src/components/HistoryDrawer.js';
import { MissionControlPage } from '../src/components/MissionControlPage.js';
import { askExcelAgent, type AgentResponse } from '../src/lib/llm-service.js';
import type { MissionRecord } from '../src/lib/missions.js';
import {
  askAgent,
  enterDemoMode,
  installBrowserStubs,
  renderApp,
  resetAppState,
} from './helpers.js';

vi.mock('../src/lib/llm-service.js', () => ({ askExcelAgent: vi.fn() }));
installBrowserStubs();

const blockedResponse: AgentResponse = {
  message: 'I could not safely prepare that request.',
  guardrail: {
    passed: false,
    schemaValid: true,
    engineValid: false,
    errors: ['The requested sheet does not exist.'],
    warnings: [],
    requiresConfirmation: false,
  },
};

const planResponse: AgentResponse = {
  message: 'Review the prepared plan.',
  plan: {
    id: 'trust-plan',
    title: 'Clean the orders export',
    description: 'Trim text and remove duplicate records before handoff.',
    assumptions: ['The first row is the header and duplicate records are identified by Order ID.'],
    steps: [
      {
        id: 'trim',
        operation: 'normalize_text',
        args: { sheet: 'Orders & Deliveries', columns: ['B'], trim: true },
        description: 'Trim the customer names in column B.',
        status: 'pending',
      },
    ],
    status: 'pending',
    totalAffectedCells: 4,
  },
};

beforeEach(() => {
  resetAppState();
  vi.mocked(askExcelAgent).mockReset();
});

describe('trust and proof surfaces', () => {
  it('explains a prepared plan with assumptions, impact, and safeguards before apply', async () => {
    enterDemoMode();
    vi.mocked(askExcelAgent).mockResolvedValue(planResponse);
    const user = userEvent.setup();
    renderApp();

    await askAgent(user, 'clean the orders export');

    expect(await screen.findByRole('region', { name: 'Plan trust summary' })).toBeInTheDocument();
    expect(screen.getByText('Why this plan')).toBeInTheDocument();
    expect(screen.getByText('Expected change')).toBeInTheDocument();
    expect(screen.getByText('4 affected cells')).toBeInTheDocument();
    expect(screen.getByText('Assumptions to review')).toBeInTheDocument();
    expect(screen.getByText(/duplicate records are identified by Order ID/i)).toBeInTheDocument();
    expect(screen.getByText('Safeguards checked')).toBeInTheDocument();
    expect(screen.getByText(/Execution is staged as one plan/i)).toBeInTheDocument();
    expect(screen.getByRole('button', { name: 'Apply All 1 Steps' })).toBeInTheDocument();
  });

  it('shows a blocked safety boundary with reasons and no apply affordance', async () => {
    enterDemoMode();
    vi.mocked(askExcelAgent).mockResolvedValue(blockedResponse);
    const user = userEvent.setup();
    const { container } = renderApp();

    await askAgent(user, 'delete a column from the missing sheet');

    expect(await screen.findByRole('alert', { name: 'Safety boundary' })).toBeInTheDocument();
    expect(screen.getByText('Safety boundary blocked this request')).toBeInTheDocument();
    expect(
      screen.getByText(
        'Nothing was changed because the requested operation did not pass the workbook safety checks.',
      ),
    ).toBeInTheDocument();
    expect(screen.getByText('Blocked: The requested sheet does not exist.')).toBeInTheDocument();
    expect(
      screen.queryByRole('button', { name: /Apply Changes|Apply All/i }),
    ).not.toBeInTheDocument();
    expect(container.textContent).toContain('Safety boundary blocked');
  });

  it('shows cell-level before/after proof in the operation audit log', async () => {
    const entry: HistoryEntry = {
      operationName: 'normalize_text',
      patch: [
        {
          kind: 'cell',
          address: { sheet: 'Orders', row: 2, column: 'A' },
          oldValue: '  North  ',
          oldType: 'string',
          newValue: 'North',
          newType: 'string',
        },
      ],
    };
    const user = userEvent.setup();
    render(
      <HistoryDrawer
        isOpen
        onClose={vi.fn()}
        entries={[entry]}
        currentPosition={1}
        onStepBack={vi.fn()}
      />,
    );

    expect(screen.getByRole('dialog', { name: 'Operation audit log' })).toBeInTheDocument();
    await user.click(screen.getByText('1 cell change'));
    expect(screen.getByText('Orders!A2')).toBeInTheDocument();
    expect(screen.getByText('North', { selector: 'strong' })).toBeInTheDocument();
    expect(screen.getByRole('listitem')).toHaveTextContent('North');
  });

  it('keeps the mission result receipt understandable and retains safety facts', () => {
    const mission: MissionRecord = {
      version: 1,
      id: 'mission-proof',
      kind: 'action',
      status: 'applied',
      title: 'Remove duplicates',
      request: 'remove duplicate rows',
      createdAt: 1,
      updatedAt: 2,
      workbook: {
        generation: 1,
        revision: 1,
        fileName: 'orders.xlsx',
        sheetName: 'Orders',
        signature: null,
      },
      safety: {
        passed: true,
        schemaValid: true,
        engineValid: true,
        requiresConfirmation: true,
        affectedCells: 1,
        errors: [],
        warnings: [],
      },
      receipt: {
        id: 'receipt-proof',
        status: 'applied',
        createdAt: 2,
        completedAt: 2,
        operations: [
          {
            name: 'delete_duplicates',
            affectedCells: 1,
            warnings: [],
            summary: 'Removed 1 duplicate row.',
          },
        ],
      },
    };
    render(
      <MissionControlPage
        missions={[mission]}
        storageStatus="saved"
        storageError=""
        unsavedIds={new Set()}
        busy={false}
        activeMissionId={null}
        onStop={vi.fn()}
        onBack={vi.fn()}
        onResume={vi.fn()}
        onReplan={vi.fn()}
        onDelete={vi.fn(async () => undefined)}
        onClear={vi.fn(async () => undefined)}
        onRetryStorage={vi.fn()}
        onDownload={vi.fn()}
      />,
    );

    expect(screen.getByText('Engine-verified commit')).toBeInTheDocument();
    expect(screen.getByText('Removed 1 duplicate row.')).toBeInTheDocument();
    expect(screen.getByRole('region', { name: 'Safety review' })).toBeInTheDocument();
    expect(screen.getByText('Confirmation')).toBeInTheDocument();
    expect(screen.getByText('required')).toBeInTheDocument();
  });
});
