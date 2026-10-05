// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createMemoryStore, sheetFingerprint } from '@excel-agent/agent';
import { createCell } from '@excel-agent/engine';

afterEach(() => {
  vi.unstubAllEnvs();
  localStorage.clear();
});

describe('runtime startup', () => {
  it('restores verified local memory without resetting outcomes or contacting a backend', async () => {
    const workbook = {
      sheets: [{ name: 'Data', rows: [[createCell('Name')], [createCell(' Ada ')]] }],
    };
    const saved = createMemoryStore();
    saved.remember({
      key: 'trim names',
      rawQuery: 'trim names',
      operation: 'normalize_text',
      sheetName: 'Data',
      args: { sheet: 'Data', columns: ['A'], trim: true },
      schemaFingerprint: sheetFingerprint(workbook, 'Data'),
    });
    for (let i = 0; i < 3; i++) saved.recordOutcome('normalize_text', 'Data', true, 'trim names');
    localStorage.setItem('excel_agent_memory_v1', saved.toJSON());
    vi.stubEnv('VITE_SUPABASE_URL', '');
    vi.stubEnv('VITE_SUPABASE_ANON_KEY', '');
    const fetchSpy = vi.spyOn(globalThis, 'fetch');
    vi.resetModules();
    try {
      const { memory, cloudMemoryConfig } = await import('./agent-runtime.js');
      expect(cloudMemoryConfig).toMatchObject({ enabled: false, url: '', apiKey: '' });
      expect(memory.entries()[0]?.successes).toBe(3);
      expect(await memory.retrieveForWorkbook('trim names', workbook, 'Data')).toMatchObject({
        operation: 'normalize_text',
        successes: 3,
      });
      expect(fetchSpy).not.toHaveBeenCalled();
    } finally {
      fetchSpy.mockRestore();
    }
  });
});
