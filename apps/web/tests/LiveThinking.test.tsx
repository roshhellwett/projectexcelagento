// @vitest-environment jsdom
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import React from 'react';
import { AgentChat, type ChatMessage } from '../src/components/AgentChat.js';
import type { SheetAudit } from '../src/lib/agent-helper.js';

describe('Live Thinking Button and Panel', () => {
  const dummyAudit: SheetAudit = {
    sheetName: 'Sheet1',
    rowCount: 10,
    columnCount: 5,
    emptyCellCount: 0,
    numericColumns: ['Revenue'],
    dateColumns: [],
    textColumns: ['Product'],
    hasMergedCells: false,
    suggestions: [],
    anomalies: [],
  };

  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('renders the Live Thinking button in the activity timeline footer', () => {
    const messages: ChatMessage[] = [
      {
        id: 'msg-1',
        sender: 'assistant',
        text: 'Inspected your data.',
        activities: [
          {
            id: 'act-1',
            type: 'planning',
            agent: 'Conductor',
            summary: 'Decomposed request into 3 segments',
            timestamp: Date.now(),
          },
        ],
        thought: 'Conductor: Analyzing user query and checking invariants...\nAll invariants passed.',
      },
    ];

    render(
      <AgentChat
        audit={dummyAudit}
        messages={messages}
        isProcessing={false}
        hasApiKey={true}
        apiKeyProvider="groq"
        onSaveApiKey={vi.fn()}
        onClearApiKey={vi.fn()}
        onSendMessage={vi.fn()}
        onApplyAction={vi.fn()}
        onUndoLast={vi.fn()}
        canUndo={false}
      />,
    );

    const thinkingBtn = screen.getByRole('button', { name: /Thinking Process/i });
    expect(thinkingBtn).toBeInTheDocument();
    expect(thinkingBtn.classList.contains('btn-live-thinking')).toBe(true);
  });

  it('shows "Live Thinking" with live pulse while processing or streaming', () => {
    const messages: ChatMessage[] = [
      {
        id: 'msg-1',
        sender: 'assistant',
        text: 'Thinking in progress...',
        isStreaming: true,
        activities: [
          {
            id: 'act-1',
            type: 'inspecting',
            agent: 'Analyst',
            summary: 'Profiling columns in parallel',
            timestamp: Date.now(),
          },
        ],
        thought: 'Data Scientist: Profiling columns A:E...',
      },
    ];

    render(
      <AgentChat
        audit={dummyAudit}
        messages={messages}
        isProcessing={true}
        hasApiKey={true}
        apiKeyProvider="groq"
        onSaveApiKey={vi.fn()}
        onClearApiKey={vi.fn()}
        onSendMessage={vi.fn()}
        onApplyAction={vi.fn()}
        onUndoLast={vi.fn()}
        canUndo={false}
      />,
    );

    const liveBtn = screen.getByRole('button', { name: /Live Thinking/i });
    expect(liveBtn).toBeInTheDocument();
    expect(liveBtn.classList.contains('is-live')).toBe(true);
  });

  it('expands the thinking panel when the button is clicked and displays written text', async () => {
    const thoughtText = 'Conductor: Formulating deterministic operation plan.\nSentinel: Verification passed.';
    const messages: ChatMessage[] = [
      {
        id: 'msg-1',
        sender: 'assistant',
        text: 'Ready to apply.',
        activities: [
          {
            id: 'act-1',
            type: 'planning',
            agent: 'Conductor',
            summary: 'Drafted plan',
            timestamp: Date.now(),
          },
        ],
        thought: thoughtText,
      },
    ];

    const { container } = render(
      <AgentChat
        audit={dummyAudit}
        messages={messages}
        isProcessing={false}
        hasApiKey={true}
        apiKeyProvider="groq"
        onSaveApiKey={vi.fn()}
        onClearApiKey={vi.fn()}
        onSendMessage={vi.fn()}
        onApplyAction={vi.fn()}
        onUndoLast={vi.fn()}
        canUndo={false}
      />,
    );

    // Initially collapsed
    expect(container.querySelector('.live-thinking-panel')).toBeNull();

    // Click button to expand
    const toggleBtn = screen.getByRole('button', { name: /Thinking Process/i });
    fireEvent.click(toggleBtn);

    // Panel is now visible
    const panel = container.querySelector('.live-thinking-panel');
    expect(panel).not.toBeNull();
    expect(screen.getByText(/Conductor: Formulating deterministic operation plan/i)).toBeInTheDocument();

    // Click again to collapse
    fireEvent.click(screen.getByRole('button', { name: /Hide Thinking/i }));
    await waitFor(() => {
      expect(container.querySelector('.live-thinking-panel')).toBeNull();
    });
  });

  it('allows clicking the status label button to toggle live thinking', () => {
    const messages: ChatMessage[] = [
      {
        id: 'msg-1',
        sender: 'assistant',
        text: 'Inspected.',
        activities: [
          {
            id: 'act-1',
            type: 'status',
            agent: 'Conductor',
            summary: 'Inspected and verified',
            timestamp: Date.now(),
          },
        ],
        thought: 'Internal reasoning details here.',
      },
    ];

    const { container } = render(
      <AgentChat
        audit={dummyAudit}
        messages={messages}
        isProcessing={false}
        hasApiKey={true}
        apiKeyProvider="groq"
        onSaveApiKey={vi.fn()}
        onClearApiKey={vi.fn()}
        onSendMessage={vi.fn()}
        onApplyAction={vi.fn()}
        onUndoLast={vi.fn()}
        canUndo={false}
      />,
    );

    const statusBtn = screen.getByRole('button', { name: /Inspected & verified/i });
    fireEvent.click(statusBtn);

    expect(container.querySelector('.live-thinking-panel')).not.toBeNull();
    expect(screen.getByText(/Internal reasoning details here/i)).toBeInTheDocument();
  });

  it('synthesizes human-readable chain-of-thought text when raw thought string is not provided', () => {
    const messages: ChatMessage[] = [
      {
        id: 'msg-1',
        sender: 'assistant',
        text: 'Result ready.',
        activities: [
          {
            id: 'act-1',
            type: 'planning',
            agent: 'Conductor',
            summary: 'Decomposed request into 2 segments',
            timestamp: Date.now(),
          },
          {
            id: 'act-2',
            type: 'inspecting',
            agent: 'Scientist',
            summary: 'Profiled 5 columns and 100 rows',
            timestamp: Date.now(),
          },
        ],
      },
    ];

    const { container } = render(
      <AgentChat
        audit={dummyAudit}
        messages={messages}
        isProcessing={false}
        hasApiKey={true}
        apiKeyProvider="groq"
        onSaveApiKey={vi.fn()}
        onClearApiKey={vi.fn()}
        onSendMessage={vi.fn()}
        onApplyAction={vi.fn()}
        onUndoLast={vi.fn()}
        canUndo={false}
      />,
    );

    fireEvent.click(screen.getByRole('button', { name: /Thinking Process/i }));

    expect(container.querySelector('.live-thinking-panel')).not.toBeNull();
    expect(screen.getByText(/\[Conductor\] Decomposed request into 2 segments/i)).toBeInTheDocument();
    expect(screen.getByText(/\[Scientist\] Profiled 5 columns and 100 rows/i)).toBeInTheDocument();
  });

  it('copies the thinking text to clipboard and updates button state', async () => {
    const writeTextMock = vi.fn().mockResolvedValue(undefined);
    Object.assign(navigator, {
      clipboard: {
        writeText: writeTextMock,
      },
    });

    const messages: ChatMessage[] = [
      {
        id: 'msg-1',
        sender: 'assistant',
        text: 'Done.',
        thought: 'Internal model reasoning to copy.',
        activities: [
          {
            id: 'act-1',
            type: 'status',
            agent: 'Conductor',
            summary: 'Done',
            timestamp: Date.now(),
          },
        ],
      },
    ];

    render(
      <AgentChat
        audit={dummyAudit}
        messages={messages}
        isProcessing={false}
        hasApiKey={true}
        apiKeyProvider="groq"
        onSaveApiKey={vi.fn()}
        onClearApiKey={vi.fn()}
        onSendMessage={vi.fn()}
        onApplyAction={vi.fn()}
        onUndoLast={vi.fn()}
        canUndo={false}
      />,
    );

    // Expand
    fireEvent.click(screen.getByRole('button', { name: /Thinking Process/i }));

    const copyBtn = screen.getByRole('button', { name: /^Copy$/i });
    expect(copyBtn).toBeInTheDocument();
    fireEvent.click(copyBtn);

    expect(writeTextMock).toHaveBeenCalledWith(
      expect.stringContaining('Internal model reasoning to copy.'),
    );
    expect(screen.getByText('Copied')).toBeInTheDocument();
  });
});
