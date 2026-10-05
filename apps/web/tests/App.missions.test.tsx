// @vitest-environment jsdom
/// <reference types="node" />
import { webcrypto } from 'node:crypto';
import { IDBFactory } from 'fake-indexeddb';
import { beforeEach, afterEach, describe, expect, it, vi } from 'vitest';
import { act, render, screen, waitFor, within } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { App } from '../src/App.js';
import { askExcelAgent, type AgentResponse } from '../src/lib/llm-service.js';
import {
  createMissionStore,
  type MissionRecord,
  type MissionStore,
  workbookSignature,
} from '../src/lib/missions.js';
import { createSampleWorkbook } from '../src/lib/engine-adapter.js';
import type { WorkspaceRecoveryStore } from '../src/lib/workspace-recovery.js';
import {
  askAgent,
  enterDemoMode,
  installBrowserStubs,
  metaPillText,
  resetAppState,
} from './helpers.js';

vi.mock('../src/lib/llm-service.js', () => ({ askExcelAgent: vi.fn() }));
installBrowserStubs();
const recovery: WorkspaceRecoveryStore = {
  load: async () => null,
  save: async () => {},
  clear: async () => {},
};
const trim: AgentResponse = {
  message: 'Ready for review.',
  proposedAction: {
    name: 'normalize_text',
    category: 'transform',
    explanation: 'Trim column B',
    args: { sheet: 'Orders & Deliveries', columns: ['B'], trim: true },
  },
};
const dedupe: AgentResponse = {
  message: 'Review duplicate removal.',
  proposedAction: {
    name: 'delete_duplicates',
    category: 'structure',
    explanation: 'Remove duplicate orders',
    args: { sheet: 'Orders & Deliveries', columns: ['A'], headerRow: 1 },
  },
};
const missionPage = () => screen.getByTestId('mission-page');
async function openMissions(user: ReturnType<typeof userEvent.setup>) {
  await user.click(screen.getByRole('button', { name: /Missions/i }));
}
async function expectStatus(store: MissionStore, status: MissionRecord['status']) {
  await waitFor(async () => expect((await store.list())[0]?.status).toBe(status));
}
beforeEach(() => {
  resetAppState();
  window.location.hash = '';
  enterDemoMode();
  vi.stubGlobal('crypto', webcrypto);
  vi.mocked(askExcelAgent).mockReset().mockResolvedValue(dedupe);
});
afterEach(() => {
  vi.unstubAllGlobals();
  window.location.hash = '';
});

describe('mission lifecycle through the real workspace controller', { timeout: 20000 }, () => {
  it('persists, remounts, rebuilds the preview, reconfirms, applies and tracks undo/redo without replaying completed work', async () => {
    const factory = new IDBFactory();
    const store = createMissionStore(() => factory);
    const user = userEvent.setup();
    const first = render(<App recoveryStore={recovery} missionRepository={store} />);
    await askAgent(user, 'remove duplicate orders');
    await screen.findByRole('button', { name: /Apply Changes/i });
    await expectStatus(store, 'prepared');
    expect((await store.list())[0]?.workbook.signature).toMatch(/^sha256:/);
    first.unmount();
    const secondStore = createMissionStore(() => factory);
    const second = render(<App recoveryStore={recovery} missionRepository={secondStore} />);
    await openMissions(user);
    await user.click(await screen.findByRole('button', { name: /Resume review/i }));
    expect(await screen.findByText(/engine generated a fresh preview/i)).toBeInTheDocument();
    await user.click(await screen.findByRole('button', { name: /Apply Changes/i }));
    await user.click(await screen.findByRole('button', { name: /Yes, apply this change/i }));
    await waitFor(() => expect(metaPillText(second.container)).toContain('10 rows'));
    await expectStatus(secondStore, 'applied');
    await user.click(screen.getByRole('button', { name: /Undo this step/i }));
    await expectStatus(secondStore, 'undone');
    await user.click(screen.getByRole('button', { name: /^Redo$/i }));
    await expectStatus(secondStore, 'applied');
    await openMissions(user);
    expect(
      within(missionPage()).queryByRole('button', { name: /Resume review/i }),
    ).not.toBeInTheDocument();
  });

  it('associates undo, redo, reset and discarded redo branches with only the committed tasks', async () => {
    const factory = new IDBFactory();
    const store = createMissionStore(() => factory);
    const user = userEvent.setup();
    render(<App recoveryStore={recovery} missionRepository={store} />);
    vi.mocked(askExcelAgent).mockResolvedValue(trim);
    await askAgent(user, 'trim column B');
    await user.click(await screen.findByRole('button', { name: /Apply Changes/i }));
    await expectStatus(store, 'applied');
    const trimId = (await store.list())[0]!.id;
    vi.mocked(askExcelAgent).mockResolvedValue(dedupe);
    await askAgent(user, 'remove duplicate orders');
    await user.click(await screen.findByRole('button', { name: /Apply Changes/i }));
    await user.click(await screen.findByRole('button', { name: /Yes, apply this change/i }));
    await waitFor(async () =>
      expect((await store.list()).filter((item) => item.status === 'applied')).toHaveLength(2),
    );
    await user.click(screen.getByRole('button', { name: /Undo this step/i }));
    await waitFor(async () => {
      const items = await store.list();
      expect(items.find((item) => item.id === trimId)?.status).toBe('applied');
      expect(items.find((item) => item.id !== trimId)?.status).toBe('undone');
    });
    expect(screen.getAllByText('Task completed')).toHaveLength(1);
    expect(screen.getAllByText('Task reverted')).toHaveLength(1);
    await user.click(screen.getByRole('button', { name: /^Redo$/i }));
    await waitFor(async () =>
      expect((await store.list()).every((item) => item.status === 'applied')).toBe(true),
    );
    await user.click(screen.getByRole('button', { name: /Reset workbook/i }));
    await waitFor(async () =>
      expect((await store.list()).every((item) => item.status === 'undone')).toBe(true),
    );
    vi.mocked(askExcelAgent).mockResolvedValue(trim);
    await askAgent(user, 'trim again on the reset workbook');
    await user.click(await screen.findByRole('button', { name: /Apply Changes/i }));
    await waitFor(async () =>
      expect((await store.list()).filter((item) => item.status === 'applied')).toHaveLength(1),
    );
    expect(screen.getByRole('button', { name: /^Redo$/i })).toBeDisabled();
    expect(screen.getAllByText('Task reverted')).toHaveLength(2);
  });

  it('marks other prepared missions stale when a different mission commits', async () => {
    const factory = new IDBFactory();
    const store = createMissionStore(() => factory);
    const user = userEvent.setup();
    const first = render(<App recoveryStore={recovery} missionRepository={store} />);
    await askAgent(user, 'remove duplicate orders');
    await screen.findByRole('button', { name: /Apply Changes/i });
    await expectStatus(store, 'prepared');
    vi.mocked(askExcelAgent).mockResolvedValue(trim);
    await askAgent(user, 'trim column B');
    const applies = await screen.findAllByRole('button', { name: /Apply Changes/i });
    await user.click(applies[1]!);
    await waitFor(async () =>
      expect((await store.list()).map((record) => record.status).sort()).toEqual([
        'applied',
        'stale',
      ]),
    );
    await openMissions(user);
    expect(screen.getByText('Needs re-check')).toBeInTheDocument();
    expect(screen.queryByRole('button', { name: /Resume review/i })).not.toBeInTheDocument();
    first.unmount();
  });

  it('displays a mismatch as needs re-check and never applies it', async () => {
    const factory = new IDBFactory();
    const store = createMissionStore(() => factory);
    const item: MissionRecord = {
      version: 1,
      id: 'stale-test',
      kind: 'action',
      status: 'prepared',
      title: 'Mismatch',
      request: 'trim column B',
      createdAt: 1,
      updatedAt: 1,
      workbook: {
        generation: 0,
        revision: 0,
        fileName: 'sample.xlsx',
        sheetName: 'Orders & Deliveries',
        signature: await workbookSignature({ ...createSampleWorkbook(), dateSystem: '1904' }),
      },
      action: trim.proposedAction,
    };
    await store.save(item);
    const user = userEvent.setup();
    const app = render(<App recoveryStore={recovery} missionRepository={store} />);
    await openMissions(user);
    await user.click(await screen.findByRole('button', { name: /Resume review/i }));
    expect(await screen.findByText('Needs re-check')).toBeInTheDocument();
    expect(
      within(missionPage()).getByText(/does not match the content inspected/i),
    ).toBeInTheDocument();
    expect(screen.queryByRole('button', { name: /Resume review/i })).not.toBeInTheDocument();
    await user.click(screen.getByRole('button', { name: /Back to workspace/i }));
    expect(metaPillText(app.container)).toContain('11 rows');
  });

  it('does not claim a failed save or delete succeeded; retry saves and removal commits before UI removal', async () => {
    const factory = new IDBFactory();
    const durable = createMissionStore(() => factory);
    let failing = true;
    const store: MissionStore = {
      ...durable,
      save: vi.fn(async (item) => {
        if (failing) throw new Error('Quota exceeded');
        await durable.save(item);
      }),
      delete: vi.fn(async (id) => {
        if (failing) throw new Error('Deletion blocked');
        await durable.delete(id);
      }),
    };
    const user = userEvent.setup();
    render(<App recoveryStore={recovery} missionRepository={store} />);
    await askAgent(user, 'remove duplicate orders');
    await screen.findByRole('button', { name: /Apply Changes/i });
    await openMissions(user);
    expect(await screen.findByText(/don’t rely on refresh recovery/i)).toBeInTheDocument();
    expect(screen.getByText('Not saved yet')).toBeInTheDocument();
    await user.click(screen.getByRole('button', { name: /^Remove$/ }));
    await user.click(screen.getByRole('button', { name: 'Delete mission data' }));
    expect(await screen.findByText(/Deletion did not commit/i)).toBeInTheDocument();
    expect(screen.getByRole('article')).toBeInTheDocument();
    await user.click(screen.getByRole('button', { name: 'Keep history' }));
    failing = false;
    await user.click(screen.getByRole('button', { name: /Retry mission storage/i }));
    await expectStatus(durable, 'prepared');
    await waitFor(() => expect(screen.queryByText('Not saved yet')).not.toBeInTheDocument());
    await user.click(screen.getByRole('button', { name: /^Remove$/ }));
    await user.click(screen.getByRole('button', { name: 'Delete mission data' }));
    await waitFor(() => expect(screen.queryByRole('article')).not.toBeInTheDocument());
    expect(await durable.list()).toEqual([]);
  });

  it('persists active planning, then cancellation, and ignores a late provider completion', async () => {
    const factory = new IDBFactory();
    const store = createMissionStore(() => factory);
    let resolve!: (value: AgentResponse) => void;
    vi.mocked(askExcelAgent).mockImplementation(
      () =>
        new Promise((done) => {
          resolve = done;
        }),
    );
    const user = userEvent.setup();
    render(<App recoveryStore={recovery} missionRepository={store} />);
    await askAgent(user, 'clean orders');
    await expectStatus(store, 'planning');
    await openMissions(user);
    await user.click(screen.getByRole('button', { name: /Stop mission/i }));
    await expectStatus(store, 'cancelled');
    await act(async () => resolve(trim));
    expect(screen.queryByRole('button', { name: /Apply Changes/i })).not.toBeInTheDocument();
    await expectStatus(store, 'cancelled');
  });

  it('turns stored active planning/execution into interrupted, not runnable, missions', async () => {
    const factory = new IDBFactory();
    const store = createMissionStore(() => factory);
    for (const status of ['planning', 'executing'] as const)
      await store.save({
        version: 1,
        id: status,
        status,
        kind: 'analysis',
        title: status,
        request: 'inspect orders',
        createdAt: 1,
        updatedAt: 1,
        workbook: {
          generation: 0,
          revision: 0,
          fileName: 'sample.xlsx',
          sheetName: 'Orders & Deliveries',
          signature: null,
        },
      });
    const user = userEvent.setup();
    render(<App recoveryStore={recovery} missionRepository={store} />);
    await openMissions(user);
    await waitFor(() => expect(screen.getAllByText('Interrupted')).toHaveLength(2));
    expect(screen.queryByRole('button', { name: /Resume review/i })).not.toBeInTheDocument();
    await waitFor(async () =>
      expect((await store.list()).every((record) => record.status === 'interrupted')).toBe(true),
    );
  });
});
