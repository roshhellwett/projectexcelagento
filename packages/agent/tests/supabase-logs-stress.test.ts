import { describe, expect, it } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { createOperationRegistry, type Workbook } from '@excel-agent/engine';
import { auditSheet, createOrchestrator } from '../src/index.js';
import { detectDelimiter, workbookFromCsvText } from '../../../apps/web/src/lib/workbook-io.js';

describe('Supabase Logs - Real-world Agent Stress Testing', () => {
  const registry = createOperationRegistry();
  const csvPath = path.resolve(__dirname, '../../../supabase_logs.csv');

  // Read actual raw user file if present, or fallback fixture
  const rawCsv = fs.existsSync(csvPath)
    ? fs.readFileSync(csvPath, 'utf-8')
    : [
        'id,date,method,pathname,status,timestamp,level,event_message,log_type,log_count,logs,auth_user',
        'row-1,"""2026-10-04T17:43:03.070Z""",,,00000,2026-10-04T17:43:03.070000,success,"statement: SET statement_timeout=\'58s\';\nCREATE FUNCTION test() RETURNS void AS $$ BEGIN NULL; END; $$;",postgres,null,[],null',
        'row-2,"""2026-10-04T17:43:01.416Z""",POST,/rest/v1/rpc,200,2026-10-04T17:43:01.416000,error,"cannot insert a non-DEFAULT value into column ""confidence""",edge,null,[],null',
      ].join('\n');

  function loadLogsWorkbook(): Workbook {
    const delimiter = detectDelimiter(rawCsv);
    expect(delimiter).toBe(',');
    const sheet = workbookFromCsvText(rawCsv, delimiter, 'Supabase Logs');
    return {
      sheets: [sheet],
      dateSystem: '1900',
    };
  }

  it('hard parsing test: accurately parses multiline SQL, preserves SQLSTATE strings, and strips outer quotes', () => {
    const wb = loadLogsWorkbook();
    const sheet = wb.sheets[0]!;

    expect(sheet.rows.length).toBeGreaterThan(1);
    const headerRow = sheet.rows[0]!.map((c) => c?.value);

    // Exactly 12 columns, correctly separated
    expect(headerRow).toEqual([
      'id',
      'date',
      'method',
      'pathname',
      'status',
      'timestamp',
      'level',
      'event_message',
      'log_type',
      'log_count',
      'logs',
      'auth_user',
    ]);

    // Row 1 checks
    const row1 = sheet.rows[1]!;
    expect(row1).toHaveLength(12);

    // Status: preserved as exact string '00000', not cast to integer 0
    expect(row1[4]?.value).toBe('00000');

    // Date: stripped of outer quotes, leaving clean ISO string
    expect(typeof row1[1]?.value).toBe('string');
    expect(row1[1]?.value).not.toMatch(/^"/);
    expect(row1[1]?.value).not.toMatch(/"$/);

    // Event message: contains multiline SQL statements with intact newlines
    const msg = String(row1[7]?.value ?? '');
    expect(msg).toContain('statement:');
    if (sheet.rows.length > 10) {
      expect(msg).toContain('CREATE OR REPLACE FUNCTION');
      expect(msg).toContain('$$;');
    }
  });

  it('hard audit test: audits 800+ complex log rows in < 150ms with clean suggestions', () => {
    const wb = loadLogsWorkbook();
    const start = performance.now();
    const audit = auditSheet(wb.sheets[0]!);
    const duration = performance.now() - start;

    expect(duration).toBeLessThan(150);
    expect(audit.sheetName).toBe('Supabase Logs');
    expect(audit.totalCols).toBe(12);
    expect(audit.totalRows).toBe(wb.sheets[0]!.rows.length);

    // Suggestions must NOT contain long un-split column lists
    for (const s of audit.suggestions) {
      expect(s.prompt).not.toContain('id,date,method');
      expect(s.prompt.length).toBeLessThan(60);
      expect(s.action.args.sheet).toBe('Supabase Logs');
    }
  });

  it('hard agent deduplication: deduplicates repetitive cron and checkpoint log entries', async () => {
    const wb = loadLogsWorkbook();
    const initialRows = wb.sheets[0]!.rows.length;
    const orchestrator = createOrchestrator({ registry });

    const decision = await orchestrator.decide({
      query: 'remove duplicate rows by event_message',
      workbook: wb,
      sheetName: 'Supabase Logs',
    });

    expect(decision.action).toBeDefined();
    expect(decision.action?.name).toBe('delete_duplicates');
    expect(decision.action?.args.columns).toEqual(['H']);
    expect(decision.guardrail?.passed).toBe(true);

    // Execute the operation
    const op = registry.get(decision.action!.name)!;
    const result = op.apply(wb, decision.action!.args);
    const finalRows = result.workbook.sheets[0]!.rows.length;

    expect(finalRows).toBeLessThan(initialRows);
    expect(finalRows).toBeGreaterThan(1);
    // Header row is preserved
    expect(result.workbook.sheets[0]!.rows[0]?.[0]?.value).toBe('id');
  });

  it('hard agent sorting: sorts 800+ log rows by timestamp chronologically', async () => {
    const wb = loadLogsWorkbook();
    const orchestrator = createOrchestrator({ registry });

    // Test ascending sort (data is initially descending in the log export)
    const decision = await orchestrator.decide({
      query: 'sort rows by timestamp ascending',
      workbook: wb,
      sheetName: 'Supabase Logs',
    });

    expect(decision.action).toBeDefined();
    expect(decision.action?.name).toBe('sort_range');
    expect(decision.action?.args.column).toBe('F'); // column F is timestamp
    expect(decision.action?.args.direction).toBe('asc');
    expect(decision.guardrail?.passed).toBe(true);

    const op = registry.get(decision.action!.name)!;
    const result = op.apply(wb, decision.action!.args);
    const rows = result.workbook.sheets[0]!.rows;

    // Header intact
    expect(rows[0]?.[5]?.value).toBe('timestamp');

    // First timestamp <= last timestamp in ascending order
    const firstTs = String(rows[1]?.[5]?.value ?? '');
    const lastTs = String(rows[rows.length - 1]?.[5]?.value ?? '');
    if (firstTs && lastTs) {
      expect(firstTs <= lastTs).toBe(true);
    }
  });

  it('hard agent text trimming: trims event_message without corrupting internal SQL newlines', async () => {
    const wb = loadLogsWorkbook();
    // Simulate real-world dirty export with leading and trailing whitespace on row 1's event_message
    const originalMsg = String(wb.sheets[0]!.rows[1]![7]?.value ?? '');
    wb.sheets[0]!.rows[1]![7] = {
      value: `   ${originalMsg}   \t`,
      type: 'string',
    };

    const orchestrator = createOrchestrator({ registry });

    const decision = await orchestrator.decide({
      query: 'trim whitespace in event_message',
      workbook: wb,
      sheetName: 'Supabase Logs',
    });

    expect(decision.action).toBeDefined();
    expect(decision.action?.name).toBe('normalize_text');
    expect(decision.action?.args.columns).toContain('H'); // column H is event_message
    expect(decision.guardrail?.passed).toBe(true);

    const op = registry.get(decision.action!.name)!;
    const result = op.apply(wb, decision.action!.args);
    const sampleMsg = String(result.workbook.sheets[0]!.rows[1]?.[7]?.value ?? '');

    // Internal newlines preserved
    expect(sampleMsg).toContain('\n');
    // Leading and trailing whitespace stripped
    expect(sampleMsg.startsWith(' ')).toBe(false);
    expect(sampleMsg.endsWith(' ')).toBe(false);
    expect(sampleMsg.endsWith('\t')).toBe(false);
  });

  it('hard agent safety & reasoning: conversational queries do not propose destructive actions', async () => {
    const wb = loadLogsWorkbook();
    const orchestrator = createOrchestrator({ registry });

    const decision = await orchestrator.decide({
      query: 'can you give me an executive summary of these database logs and any errors?',
      workbook: wb,
      sheetName: 'Supabase Logs',
    });

    // Pure conversational overview: no unguardrailed mutations
    expect(decision.action).toBeUndefined();
    expect(decision.message).toBeTruthy();
  });

  it('hard agent filtering: filters error logs cleanly preserving JSON payloads and stack traces', async () => {
    const wb = loadLogsWorkbook();
    const orchestrator = createOrchestrator({ registry });

    const decision = await orchestrator.decide({
      query: 'filter rows where level is error',
      workbook: wb,
      sheetName: 'Supabase Logs',
    });

    expect(decision.action).toBeDefined();
    expect(decision.action?.name).toBe('filter_rows');
    expect(decision.action?.args.column).toBe('G'); // column G is level
    expect(decision.action?.args.operator).toBe('equals');
    expect(decision.action?.args.value).toBe('error');
    expect(decision.guardrail?.passed).toBe(true);

    const op = registry.get(decision.action!.name)!;
    const result = op.apply(wb, decision.action!.args);
    const filteredRows = result.workbook.sheets[0]!.rows;

    // Header preserved
    expect(filteredRows[0]?.[6]?.value).toBe('level');
    // Only error rows remain
    for (let i = 1; i < filteredRows.length; i++) {
      expect(filteredRows[i]?.[6]?.value).toBe('error');
    }
  });

  it('hard dedupe audit: gracefully reports 0 duplicates when entire row contains globally unique UUIDs', async () => {
    const wb = loadLogsWorkbook();
    const orchestrator = createOrchestrator({ registry });

    // When deduplicating across ALL columns, all 836 rows have unique UUIDs in column A
    const decision = await orchestrator.decide({
      query: 'remove duplicate rows from these logs',
      workbook: wb,
      sheetName: 'Supabase Logs',
    });

    // Clean informative message rather than a broken 0-cell mutation
    expect(decision.action).toBeUndefined();
    expect(decision.message).toContain('0 duplicate row');
  });

  it('hard formula injection and SQL-like safety: multiline SQL and symbols never execute as formulas', () => {
    const wb = loadLogsWorkbook();
    const sheet = wb.sheets[0]!;

    // Scan all event_message and pathname cells for potential formula injection triggers
    for (let r = 1; r < Math.min(sheet.rows.length, 50); r++) {
      const msgCell = sheet.rows[r]?.[7];
      if (msgCell && typeof msgCell.value === 'string') {
        // Value must be string type, never mistakenly parsed as formula
        expect(msgCell.formula).toBeUndefined();
        expect(msgCell.type).toBe('string');
      }
    }
  });
});
