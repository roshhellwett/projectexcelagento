import { describe, expect, it } from 'vitest';
import { createCell, createOperationRegistry, type Workbook } from '@excel-agent/engine';
import { evidenceFromToolResult } from '../src/evidence.js';
import { createOrchestrator } from '../src/index.js';

const workbook: Workbook = {
  sheets: [
    {
      name: 'Sales',
      rows: [
        [createCell('Region'), createCell('Revenue')],
        [createCell('North'), createCell(100)],
        [createCell('South'), createCell(200)],
      ],
    },
  ],
};

describe('evidence provenance', () => {
  it('turns deterministic statistics into source-addressed facts', () => {
    const item = evidenceFromToolResult(
      'describe_column',
      { sheet: 'Sales' },
      {
        sheet: 'Sales',
        column: 'B',
        headerName: 'Revenue',
        headerRow: 1,
        totalCount: 2,
        numericCount: 2,
        missingCount: 0,
        nonNumericCount: 0,
        mean: 150,
        median: 150,
        outlierCount: 0,
      },
    );
    expect(item).toMatchObject({ kind: 'statistic', source: 'Sales!B2:B3' });
    expect(item?.facts).toEqual(
      expect.arrayContaining([
        { label: 'Mean', value: '150' },
        { label: 'Numeric observations', value: '2' },
      ]),
    );
  });

  it('does not convert errors or external prose into workbook evidence', () => {
    expect(
      evidenceFromToolResult('describe_column', { sheet: 'Sales' }, { error: 'Column not found' }),
    ).toBeUndefined();
    expect(
      evidenceFromToolResult('search_web', {}, { title: 'External page', snippet: 'text' }),
    ).toBeUndefined();
  });

  it('attaches evidence to the no-key statistical analyst path', async () => {
    const decision = await createOrchestrator({ registry: createOperationRegistry() }).decide({
      query: 'descriptive statistics for column B',
      workbook,
      sheetName: 'Sales',
    });
    expect(decision.evidence).toHaveLength(1);
    expect(decision.evidence?.[0]).toMatchObject({ kind: 'statistic', source: 'Sales!B2:B3' });
  });
});
