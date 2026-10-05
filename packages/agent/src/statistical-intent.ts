import type { Workbook } from '@excel-agent/engine';
import type { ColumnMetadata } from './analysis.js';
import type { ClarificationQuestion } from './types.js';
import { analyzeColumnRelationship, describeColumn } from './statistical-tools.js';

function display(value: number | null): string {
  return value === null ? 'Undefined' : String(Number(value.toPrecision(10)));
}

export function analyzeStatisticalIntent(query: string, workbook: Workbook, sheetName: string, columns: ColumnMetadata[]): { message: string; clarification?: ClarificationQuestion } | undefined {
  if (/\b(?:formula|syntax|how\s+(?:to|do))\b/i.test(query)) return undefined;
  const isRelationship = /\b(?:correlat(?:ion|e)|pearson|linear\s+regression|regress)\b/i.test(query);
  if (!isRelationship && !/\b(?:descriptive\s+statistics|summary\s+statistics|statistical\s+summary|statistics|median|standard\s+deviation|stdev|variance|quartiles?|outliers?)\b/i.test(query)) return undefined;

  const explicit: ColumnMetadata[] = [];
  for (const match of query.matchAll(/\b(?:columns?|col)\s+([a-z]{1,3})\b(?:\s*(?:and|,|vs\.?)\s*(?:columns?|col)?\s*([a-z]{1,3})\b)?/gi)) {
    for (const letter of [match[1], match[2]]) {
      if (!letter) continue;
      const column = columns.find((column) => column.letter.toLowerCase() === letter.toLowerCase())
        ?? columns.find((column) => column.cleanName === letter.toLowerCase());
      if (!column) return { message: `Column **${letter.toUpperCase()}** does not exist in **${sheetName}**.` };
      if (!explicit.includes(column)) explicit.push(column);
    }
  }
  const mentioned = explicit.length > 0 ? explicit : columns
    .filter((column) => column.rawName && new RegExp(`\\b${column.rawName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`, 'i').test(query))
    .sort((a, b) => query.toLowerCase().indexOf(a.cleanName) - query.toLowerCase().indexOf(b.cleanName));

  if (isRelationship) {
    if (mentioned.length !== 2) return { message: 'Specify two columns, for example **"correlation between column A and column B"** or **"linear regression of column B on column A"** (B is the response, A the predictor).' };
    const regressOn = /\b(?:regression\s+(?:of\s+)?|regress\s+).+\s+on\s+/i.test(query);
    const x = mentioned[regressOn ? 1 : 0]!;
    const y = mentioned[regressOn ? 0 : 1]!;
    const result = analyzeColumnRelationship(workbook, sheetName, x.letter, y.letter);
    if ('error' in result) return { message: result.error! };
    return { message: `### ${/regress/i.test(query) ? 'Linear regression' : 'Correlation'}: ${y.rawName} vs ${x.rawName}\n\n` +
      `- **Complete numeric pairs:** ${result.pairCount}\n- **Excluded rows:** ${result.excludedPairCount}\n` +
      `- **Pearson correlation (r):** ${display(result.correlation)}\n- **R²:** ${display(result.rSquared)}\n` +
      `- **Slope:** ${display(result.slope)}\n- **Intercept:** ${display(result.intercept)}\n` +
      `- **Residual standard error:** ${display(result.residualStandardError)}\n\n` +
      (result.slope !== null && result.intercept !== null ? `Fitted equation: **${y.rawName} = ${display(result.intercept)} + ${display(result.slope)} × ${x.rawName}**.\n\n` : '') +
      `${result.warnings.join(' ')}\n\nComputed from all complete numeric pairs using ordinary least squares with an intercept. Correlation measures association; it does not establish causation.` };
  }

  const numericColumns = columns.filter((column) => column.numericValues.length > 0);
  const target = mentioned.length === 1 ? mentioned[0] : mentioned.length === 0 && numericColumns.length === 1 ? numericColumns[0] : undefined;
  if (!target) return {
    message: 'Which column should I analyze? Choose a column for descriptive statistics and outlier detection.',
    clarification: { question: 'Select a statistics column:', options: columns.slice(0, 12).map((column) => ({
      label: `${column.rawName} (${column.letter})`, query: `descriptive statistics for column ${column.letter}`,
    })) },
  };
  const result = describeColumn(workbook, sheetName, target.letter);
  if ('error' in result) return { message: result.error! };
  const metrics: [string, number | null][] = [
    ['Sum', result.sum], ['Mean', result.mean], ['Median', result.median], ['Minimum', result.min], ['Maximum', result.max],
    ['Q1 (25th percentile)', result.q1], ['Q3 (75th percentile)', result.q3],
    ['Sample standard deviation', result.sampleStandardDeviation], ['Sample variance', result.sampleVariance],
    ['Population standard deviation', result.populationStandardDeviation], ['IQR', result.iqr],
  ];
  return { message: `### Descriptive statistics: ${target.rawName} (${target.letter})\n\n` +
    `**${result.numericCount} numeric observations** across ${result.totalCount} data rows; ${result.missingCount} missing and ${result.nonNumericCount} nonnumeric values excluded.\n\n` +
    `| Statistic | Value |\n| --- | ---: |\n${metrics.map(([label, value]) => `| ${label} | ${display(value)} |`).join('\n')}\n\n` +
    `**Outliers (1.5 × IQR): ${result.outlierCount}**. Fences: ${display(result.lowerFence)} to ${display(result.upperFence)}.\n` +
    result.outliers.map((outlier) => `- Row ${outlier.rowNumber}: ${display(outlier.value)}`).join('\n') +
    (result.outlierCount > result.outliers.length ? '\nShowing the first 25 outliers.' : '') +
    `\n\n${result.warnings.join(' ')}\nInclusive quartiles match Excel PERCENTILE.INC. Results use the full column; outlier flags do not remove data.` };
}
