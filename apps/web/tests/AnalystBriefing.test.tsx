// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { render, screen, within } from '@testing-library/react';
import { createCell, type Workbook } from '@excel-agent/engine';
import { AnalystBriefing } from '../src/components/AnalystBriefing.js';

function workbook(values: (string | number | boolean | Date | null)[][]): Workbook {
  return {
    sheets: [{ name: 'Data', rows: values.map((row) => row.map((value) => createCell(value))) }],
  };
}

function numericRow(header: string) {
  const table = screen.getByRole('table', { name: /Full-column descriptive statistics/ });
  const row = within(table)
    .getByRole('rowheader', { name: new RegExp(header) })
    .closest('tr')!;
  return within(row).getAllByRole('cell');
}

describe('AnalystBriefing', () => {
  it('renders full-range negative and blank statistics and source evidence without providers or a baseline', () => {
    const wb = workbook([
      ['Amount', 'Segment'],
      [-10, 'A'],
      [-5, 'B'],
      [null, 'A'],
      ['-2', null],
      [],
      ['invalid', 'B'],
    ]);
    render(<AnalystBriefing workbook={wb} sheetName="Data" />);
    expect(screen.getByRole('heading', { name: 'Analyst briefing' })).toBeInTheDocument();
    const source = screen.getByText('FULL STORED RANGE').parentElement!;
    expect(within(source).getByText("'Data'!A1:B7")).toBeInTheDocument();
    const cells = numericRow('Amount');
    expect(cells.map((cell) => cell.textContent).slice(0, 7)).toEqual([
      '6',
      '3',
      '2',
      '1',
      '-17',
      '-5.6666667',
      '-5',
    ]);
    expect(screen.getByText(/No baseline supplied/)).toBeInTheDocument();
    expect(screen.getByText(/Presence only—not correctness/)).toBeInTheDocument();
    expect(screen.getByText(/No model required/)).toBeInTheDocument();
    expect(screen.getByRole('region', { name: 'Methods and limits' })).toHaveTextContent(
      'no causal, financial-performance, or forecasting claims',
    );
    expect(screen.getByText("'Data'!A4:A4")).toBeInTheDocument();
  });

  it('shows live formula results and excluded errors/unsupported formulas rather than cached magnitudes', () => {
    const wb = workbook([['Result'], [999], [999], [999]]);
    wb.sheets[0]!.rows[1]![0] = createCell(999, { formula: '=2*3' });
    wb.sheets[0]!.rows[2]![0] = createCell(999, { formula: '=1/0' });
    wb.sheets[0]!.rows[3]![0] = createCell(999, { formula: '=UNSUPPORTED()' });
    render(<AnalystBriefing workbook={wb} sheetName="Data" />);
    const cells = numericRow('Result');
    expect(cells.slice(0, 6).map((cell) => cell.textContent)).toEqual([
      '3',
      '1',
      '0',
      '2',
      '6',
      '6',
    ]);
    expect(
      screen.getByRole('table', { name: /Full-column descriptive statistics/ }),
    ).not.toHaveTextContent('999');
    const quality = screen.getByRole('region', { name: 'Data quality summary' });
    expect(quality).toHaveTextContent('Error values excluded1');
    expect(quality).toHaveTextContent('Unsupported formulas / dependents1');
    expect(screen.getByText('Error value #DIV/0!; excluded')).toBeInTheDocument();
  });

  it('provides named SVG charts with exact visible table alternatives and bounded categories', () => {
    const wb = workbook([
      ['Segment'],
      ...Array.from({ length: 19 }, (_, index) => [`Group ${String(index).padStart(2, '0')}`]),
      ['Group 18'],
    ]);
    render(<AnalystBriefing workbook={wb} sheetName="Data" />);
    const chart = screen.getByRole('img', { name: /Segment \(A\) categorical counts/ });
    expect(chart).toHaveAccessibleName('Segment (A) categorical counts');
    expect(chart).toHaveAccessibleDescription(/Exact labels and counts are in the table below/);
    const table = screen.getByRole('table', { name: 'Segment (A) chart data · exact counts' });
    expect(within(table).getAllByRole('row')).toHaveLength(9);
    const top = within(table).getByRole('rowheader', { name: 'Group 18 (string)' }).closest('tr')!;
    expect(
      within(top)
        .getAllByRole('cell')
        .map((cell) => cell.textContent),
    ).toEqual(['1', '2']);
    expect(table).not.toHaveTextContent('Group 17');
    expect(screen.getByText('9 observations shown · 11 omitted.')).toBeInTheDocument();
    expect(screen.getByText(/Top 8 of 19 exact/)).toBeInTheDocument();
    expect(screen.getAllByText("'Data'!A2:A21").length).toBeGreaterThan(0);
  });

  it('handles constant data, no numeric observations and empty/missing sheets honestly', () => {
    const { rerender } = render(
      <AnalystBriefing
        workbook={workbook([
          ['Constant', 'Words'],
          [5, 'alpha'],
          [5, 'beta'],
        ])}
        sheetName="Data"
      />,
    );
    const table = screen.getByRole('table', { name: 'Constant (A) chart data · exact counts' });
    expect(within(table).getByRole('rowheader', { name: '5' })).toBeInTheDocument();
    expect(numericRow('Constant')[11]).toHaveTextContent('0');
    expect(numericRow('Words')[4]).toHaveTextContent('Undefined');
    rerender(
      <AnalystBriefing workbook={workbook([['Words'], ['alpha'], [null]])} sheetName="Data" />,
    );
    expect(
      screen.getByText(/No numeric observations; numeric totals and distributions are undefined/),
    ).toBeInTheDocument();
    rerender(<AnalystBriefing workbook={workbook([['Header']])} sheetName="Data" />);
    expect(
      screen.getByText(/No chartable numeric or categorical observations/),
    ).toBeInTheDocument();
    expect(numericRow('Header')[4]).toHaveTextContent('Undefined');
    rerender(<AnalystBriefing workbook={workbook([])} sheetName="Missing" />);
    expect(screen.getByRole('status')).toHaveTextContent('Sheet “Missing” was not found');
    expect(
      screen.queryByRole('table', { name: /Full-column descriptive statistics/ }),
    ).not.toBeInTheDocument();
  });

  it('renders shifted-header baseline as positional evidence, not comparable financial results', () => {
    const baseline = workbook([
      ['Revenue', 'Cost'],
      [100, 10],
    ]);
    const current = workbook([
      ['Cost', 'Revenue'],
      [10, 100],
    ]);
    render(<AnalystBriefing workbook={current} sheetName="Data" baseline={baseline} />);
    const comparison = screen.getByRole('region', { name: 'Baseline comparison' });
    expect(comparison).toHaveTextContent('4 stored cells changed');
    expect(comparison).toHaveTextContent('Changed—positional only');
    expect(comparison).toHaveTextContent('Header changed');
    expect(comparison).toHaveTextContent(
      'No record alignment, trends, or comparable aggregate deltas are inferred',
    );
    expect(comparison).toHaveTextContent('Layouts differ');
    const table = within(comparison).getByRole('table', { name: /Exact stored-cell examples/ });
    expect(within(table).getAllByRole('row')).toHaveLength(5);
    expect(table).toHaveTextContent('"value": "Revenue"');
    expect(table).toHaveTextContent('"value": "Cost"');
  });

  it('renders exact formula and Date change evidence plus the workbook epoch warning', () => {
    const baseline = workbook([
      ['Date', 'Formula'],
      [new Date('2024-01-01T00:00:00Z'), 4],
    ]);
    baseline.sheets[0]!.rows[1]![1] = createCell(4, { formula: '=2+2' });
    const current = structuredClone(baseline);
    current.dateSystem = '1904';
    current.sheets[0]!.rows[1]![0] = createCell(new Date('2024-01-02T00:00:00Z'));
    current.sheets[0]!.rows[1]![1] = createCell(4, { formula: '=1+3' });
    render(<AnalystBriefing workbook={current} sheetName="Data" baseline={baseline} />);
    const comparison = screen.getByRole('region', { name: 'Baseline comparison' });
    expect(comparison).toHaveTextContent('Date epoch: 1900 → 1904 (changed)');
    expect(comparison).toHaveTextContent('The workbook date epoch changed');
    const evidence = within(comparison).getByRole('table', { name: /Exact stored-cell examples/ });
    expect(evidence).toHaveTextContent('2024-01-01T00:00:00.000Z');
    expect(evidence).toHaveTextContent('"epochMilliseconds": 1704067200000');
    expect(evidence).toHaveTextContent('2024-01-02T00:00:00.000Z');
    expect(evidence).toHaveTextContent('"formula": "=2+2"');
    expect(evidence).toHaveTextContent('"formula": "=1+3"');
  });

  it('updates from immutable props and bounds baseline evidence without understating full change counts', () => {
    const baseline = workbook([['Number'], ...Array.from({ length: 50 }, (_, index) => [index])]);
    const current = workbook([
      ['Number'],
      ...Array.from({ length: 50 }, (_, index) => [index + 1]),
    ]);
    const { rerender } = render(
      <AnalystBriefing workbook={baseline} sheetName="Data" baseline={baseline} />,
    );
    expect(screen.getByText(/No exact stored-cell changes/)).toBeInTheDocument();
    rerender(<AnalystBriefing workbook={current} sheetName="Data" baseline={baseline} />);
    const comparison = screen.getByRole('region', { name: 'Baseline comparison' });
    expect(comparison).toHaveTextContent('50 stored cells changed');
    expect(comparison).toHaveTextContent('20 examples · 30 additional changes omitted');
    expect(
      within(
        within(comparison).getByRole('table', { name: /Exact stored-cell examples/ }),
      ).getAllByRole('row'),
    ).toHaveLength(21);
    expect(numericRow('Number')[4]).toHaveTextContent('1,275');
  });
});
