import type { EvidenceItem } from './types.js';

const display = (value: unknown): string => {
  if (value === null || value === undefined || value === '') return '—';
  if (typeof value === 'number')
    return Number.isFinite(value)
      ? value.toLocaleString(undefined, { maximumFractionDigits: 6 })
      : 'Undefined';
  return String(value);
};

const sourceFor = (sheet: unknown, column: unknown, start = 2, end?: unknown): string => {
  const name = String(sheet ?? 'Active sheet');
  const col = String(column ?? '').trim();
  if (!col) return name;
  return `${name}!${col}${start}${end === undefined ? '' : `:${col}${end}`}`;
};

export function evidenceFromToolResult(
  name: string,
  args: Record<string, unknown>,
  output: unknown,
): EvidenceItem | undefined {
  if (!output || typeof output !== 'object' || Array.isArray(output)) return undefined;
  const result = output as Record<string, unknown>;
  if (typeof result.error === 'string') return undefined;
  const sheet = result.sheet ?? args.sheet ?? 'Active sheet';

  if (name === 'describe_column') {
    const facts = [
      ['Numeric observations', result.numericCount],
      ['Data rows', result.totalCount],
      ['Missing', result.missingCount],
      ['Nonnumeric', result.nonNumericCount],
      ['Mean', result.mean],
      ['Median', result.median],
      ['IQR outliers', result.outlierCount],
    ]
      .filter(([, value]) => value !== undefined)
      .map(([label, value]) => ({ label: String(label), value: display(value) }));
    return {
      id: `evidence-${name}-${String(result.column)}`,
      kind: 'statistic',
      title: `Full-column statistics · ${String(result.headerName ?? result.column)}`,
      source: sourceFor(
        sheet,
        result.column,
        Number(result.headerRow ?? 1) + 1,
        Number(result.totalCount ?? 0) + Number(result.headerRow ?? 1),
      ),
      facts,
      note: 'Computed from the full column with supported formula values evaluated live.',
    };
  }

  if (name === 'analyze_column_relationship') {
    const facts = [
      ['Complete numeric pairs', result.pairCount],
      ['Excluded rows', result.excludedPairCount],
      ['Pearson r', result.correlation],
      ['R²', result.rSquared],
      ['Slope', result.slope],
      ['Intercept', result.intercept],
    ]
      .filter(([, value]) => value !== undefined)
      .map(([label, value]) => ({ label: String(label), value: display(value) }));
    return {
      id: `evidence-${name}-${String(result.xColumn)}-${String(result.yColumn)}`,
      kind: 'statistic',
      title: `Relationship analysis · ${String(result.yHeader ?? result.yColumn)} vs ${String(result.xHeader ?? result.xColumn)}`,
      source: `${String(sheet)}!${String(result.xColumn)} ↔ ${String(result.yColumn)}`,
      facts,
      note: 'Pearson correlation and ordinary least squares; association is not causation.',
    };
  }

  if (name === 'calculate_aggregate') {
    return {
      id: `evidence-${name}-${String(result.column)}`,
      kind: 'aggregate',
      title: `${String(result.metric)} · ${String(result.column)}`,
      source: sourceFor(sheet, result.column),
      facts: [
        { label: 'Result', value: display(result.value) },
        { label: 'Numeric rows', value: display(result.count) },
      ],
    };
  }

  if (name === 'profile_column') {
    return {
      id: `evidence-${name}-${String(result.column)}`,
      kind: 'profile',
      title: `Column profile · ${String(result.headerName ?? result.column)}`,
      source: sourceFor(sheet, result.column),
      facts: [
        ['Type', result.inferredType],
        ['Nonblank', result.nonBlankCount],
        ['Blank', result.blankCount],
        ['Distinct', result.distinctCount],
      ]
        .filter(([, value]) => value !== undefined)
        .map(([label, value]) => ({ label: String(label), value: display(value) })),
    };
  }

  if (name === 'search_sheet') {
    return {
      id: `evidence-${name}-${String(result.query)}`,
      kind: 'inspection',
      title: `Search evidence · “${String(result.query ?? '')}”`,
      source: String(sheet),
      facts: [
        { label: 'Matches', value: display(result.totalMatches) },
        {
          label: 'Returned samples',
          value: display(Array.isArray(result.matches) ? result.matches.length : 0),
        },
      ],
    };
  }

  if (name === 'read_cell_range') {
    return {
      id: `evidence-${name}-${String(sheet)}-${String(result.startRow)}`,
      kind: 'inspection',
      title: 'Range inspected',
      source: `${String(sheet)}!${String(result.startColumn)}${String(result.startRow)}:${String(result.endColumn)}${String(result.endRow)}`,
      facts: [
        { label: 'Rows read', value: display(Array.isArray(result.rows) ? result.rows.length : 0) },
      ],
    };
  }

  return undefined;
}
