import { describe, expect, it, vi } from 'vitest';
import { applyOperation, createCell, createOperationRegistry, type Workbook } from '@excel-agent/engine';
import { analyzeColumnRelationship, analyzeSpreadsheetIntentAndData, createOrchestrator, describeColumn } from '../src/index.js';

function workbook(): Workbook {
  return { sheets: [{ name: 'Data', rows: [
    [createCell('Spend'), createCell('Revenue')],
    ...[1, 2, 3, 4, 100].map((x) => [createCell(x), createCell(2 + 3 * x)]),
    [createCell(null), createCell(20)],
  ] }] };
}

describe('statistical read tools', () => {
  it('identifies full-column outliers with original row numbers and leaves the sheet intact', () => {
    const before = workbook();
    expect(describeColumn(before, 'Data', 'A')).toMatchObject({ numericCount: 5, missingCount: 1, median: 3, outlierCount: 1, outliers: [{ rowNumber: 6, value: 100 }] });
    expect(before).toEqual(workbook());
  });

  it('fits complete numeric pairs and refuses invalid references', () => {
    expect(analyzeColumnRelationship(workbook(), 'Data', 'A', 'B')).toMatchObject({ pairCount: 5, excludedPairCount: 1, slope: 3, intercept: 2 });
    expect(analyzeColumnRelationship(workbook(), 'Data', 'A', 'A')).toHaveProperty('error');
    expect(describeColumn(workbook(), 'Data', 'ZZ')).toHaveProperty('error');
    expect(describeColumn(workbook(), 'Missing', 'A')).toHaveProperty('error');
    expect(describeColumn(workbook(), 'Data', 'A', 100)).toHaveProperty('error');
  });

  it('uses live formula results for both read tools and engine aggregates', () => {
    const wb: Workbook = { sheets: [{ name: 'Data', rows: [
      [createCell('X'), createCell('Y')],
      [createCell(2), createCell(999, { formula: 'A2*3' })],
      [createCell(4), createCell(999, { formula: 'A3*3' })],
    ] }] };
    expect(describeColumn(wb, 'Data', 'B')).toMatchObject({ mean: 9, sum: 18 });
    expect(analyzeColumnRelationship(wb, 'Data', 'A', 'B')).toMatchObject({ slope: 3, intercept: 0 });
    const result = applyOperation(wb, 'aggregate_column', { sheet: 'Data', column: 'B', aggregation: 'sum' });
    expect(result.ok && result.report.aggregate).toBe(18);
  });

  it('supports non-first-row headers through the read tool contract', () => {
    const wb = workbook();
    wb.sheets[0]!.rows.unshift([createCell('Report title')]);
    expect(describeColumn(wb, 'Data', 'A', 2)).toMatchObject({ headerName: 'Spend', numericCount: 5, outliers: [{ rowNumber: 7, value: 100 }] });
  });
});

describe('offline statistical requests', () => {
  it.each(['descriptive statistics for column A', 'median of Spend', 'standard deviation of Spend', 'find outliers in column A'])('answers %s directly without proposing a mutation', (query) => {
    const result = analyzeSpreadsheetIntentAndData(query, workbook(), 'Data');
    expect(result.proposedAction).toBeUndefined();
    expect(result.message).toContain('Descriptive statistics: Spend (A)');
    expect(result.message).toContain('Row 6: 100');
  });

  it('honors response/predictor ordering for regression and named columns for correlation', () => {
    expect(analyzeSpreadsheetIntentAndData('linear regression of column B on column A', workbook(), 'Data').message).toContain('Revenue = 2 + 3 × Spend');
    expect(analyzeSpreadsheetIntentAndData('correlation between Spend and Revenue', workbook(), 'Data').message).toContain('**Pearson correlation (r):** 1');
  });

  it('asks for a column when the request is ambiguous', () => {
    expect(analyzeSpreadsheetIntentAndData('descriptive statistics', workbook(), 'Data').clarification).toBeDefined();
    expect(analyzeSpreadsheetIntentAndData('correlation', workbook(), 'Data').message).toContain('Specify two columns');
  });

  it('routes a statistical request through the no-key orchestrator', async () => {
    const decision = await createOrchestrator({ registry: createOperationRegistry() }).decide({ query: 'median of Spend', workbook: workbook(), sheetName: 'Data' });
    expect(decision.action).toBeUndefined();
    expect(decision.message).toContain('| Median | 3 |');
    expect(decision.telemetry).toBeUndefined();
  });

  it.each([
    ['describe_column', { column: 'A' }, { median: 3, outlierCount: 1 }],
    ['analyze_column_relationship', { xColumn: 'A', yColumn: 'B' }, { pairCount: 5, slope: 3 }],
  ])('executes model-requested %s against the real workbook', async (name, args, expected) => {
    let calls = 0;
    let toolResult: unknown;
    const fetchSpy = vi.spyOn(globalThis, 'fetch').mockImplementation(async (_url, init) => {
      const body = JSON.parse(String(init?.body)) as { messages: { role: string; content: string }[] };
      calls += 1;
      if (calls > 1) {
        const content = body.messages.find((message) => message.role === 'tool')!.content;
        toolResult = JSON.parse(content.slice(content.indexOf('\n') + 1));
      }
      return new Response(JSON.stringify({ model: 'llama-3.3-70b-versatile', choices: [{ index: 0,
        message: calls === 1 ? { role: 'assistant', content: '', tool_calls: [{ id: 'read-1', type: 'function', function: { name, arguments: JSON.stringify(args) } }] } : { role: 'assistant', content: 'Here are the computed results.' },
        finish_reason: calls === 1 ? 'tool_calls' : 'stop',
      }] }), { headers: { 'content-type': 'application/json' } });
    });
    try {
      const decision = await createOrchestrator({ registry: createOperationRegistry() }).decide({ query: 'inspect dependence for me', workbook: workbook(), sheetName: 'Data', config: { provider: 'groq', apiKey: 'gsk_test' } });
      expect(toolResult, JSON.stringify({ calls, decision })).toMatchObject(expected);
      expect(decision.action).toBeUndefined();
      expect(decision.message).toBe('Here are the computed results.');
    } finally {
      fetchSpy.mockRestore();
    }
  });
});
