import {
  columnToIndex,
  indexToColumn,
  maxColumnCount,
  type Sheet,
  type Workbook,
} from '@excel-agent/engine';
import type { ProposedAction } from './analysis.js';
import { isComplexRequest } from './complexity.js';
import type { ClarificationQuestion, EvidenceItem } from './types.js';

export interface ReconciliationIntent {
  message: string;
  proposedAction?: ProposedAction;
  clarification?: ClarificationQuestion;
  evidence?: EvidenceItem[];
}
export const isReconciliationRequest = (query: string) =>
  /\b(?:reconcile|reconciliation|reconcile_sheets)\b/i.test(query);
const normalized = (value: string) =>
  value
    .trim()
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]/gu, '');
const quote = (value: string) => JSON.stringify(value);
const escape = (value: string) => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const last = (query: string, pattern: RegExp) => [...query.matchAll(pattern)].at(-1)?.[1]?.trim();
const decode = (value: string) => {
  try {
    return value.startsWith('"') ? String(JSON.parse(value)) : value.trim();
  } catch {
    return value.trim();
  }
};
const field = (query: string, name: string) =>
  last(query, new RegExp(`\\b${name}\\s*:\\s*([^;]+)`, 'gi'));
interface Column {
  column: string;
  header: string;
  name: string;
}
function columns(sheet: Sheet, headerRow: number): Column[] {
  return Array.from({ length: maxColumnCount(sheet.rows) }, (_, index) => {
    const cell = sheet.rows[headerRow - 1]?.[index];
    const header =
      cell && cell.formula === undefined && typeof cell.value === 'string' ? cell.value : '';
    return { column: indexToColumn(index), header, name: normalized(header) };
  });
}
function resolveColumn(value: string, available: Column[]): string | undefined {
  const raw = value.replace(/^column\s+/i, '').trim();
  const name = decode(raw);
  const explicitIndex =
    !raw.startsWith('"') && /^[A-Z]{1,3}$/.test(raw) ? columnToIndex(raw) : undefined;
  if (explicitIndex !== undefined && explicitIndex < available.length)
    return indexToColumn(explicitIndex);
  const byName = available.filter((column) => column.name && column.name === normalized(name));
  if (byName.length === 1) return byName[0]!.column;
  if (byName.length > 1) return undefined;
  const index = columnToIndex(name);
  return index !== undefined && index < available.length ? indexToColumn(index) : undefined;
}
function resolveKeys(value: string, available: Column[]): string[] | undefined {
  // A header containing commas or "and" is first treated as one exact named column.
  const single =
    !/\s*(?:,|\band\b|&)\s*/i.test(value) || (value.startsWith('"') && value.endsWith('"'))
      ? resolveColumn(value, available)
      : undefined;
  if (single) return [single];
  const parts = value.split(/\s*(?:,|\band\b|&)\s*/i).filter(Boolean);
  const resolved = parts.map((part) => resolveColumn(part, available));
  return parts.length >= 1 &&
    parts.length <= 5 &&
    resolved.every(Boolean) &&
    new Set(resolved).size === resolved.length
    ? (resolved as string[])
    : undefined;
}
const columnLabel = (column: Column) => `${column.header || 'Unnamed column'} (${column.column})`;
function question(
  message: string,
  prompt: string,
  options: ClarificationQuestion['options'] = [],
): ReconciliationIntent {
  return { message, clarification: { question: prompt, options: options.slice(0, 8) } };
}
const append = (query: string, setting: string) => `${query}; ${setting}`;
function sheetMentions(query: string, sheets: Sheet[]): Sheet[] {
  const occupied: { start: number; end: number; sheet: Sheet }[] = [];
  for (const sheet of [...sheets].sort((a, b) => b.name.length - a.name.length)) {
    for (const match of query.matchAll(
      new RegExp(`(?<![\\p{L}\\p{N}])${escape(sheet.name)}(?![\\p{L}\\p{N}])`, 'giu'),
    )) {
      const start = match.index;
      const end = start + match[0].length;
      if (!occupied.some((item) => start < item.end && end > item.start))
        occupied.push({ start, end, sheet });
    }
  }
  return occupied
    .sort((a, b) => a.start - b.start)
    .map((item) => item.sheet)
    .filter((sheet, index, all) => all.indexOf(sheet) === index);
}
const identifier = (column: Column) =>
  /^(?:id|key|reference|ref|sku|code|(?:order|invoice|transaction|record|document|account|customer|product|item|purchaseorder|po)(?:id|number|no|code|reference))$/.test(
    column.name,
  );
const amount = (column: Column) =>
  /^(?:amount|total|balance|value|netamount|grossamount|invoicetotal|invoiceamount|ordertotal|orderamount|transactionamount|revenue|salesamount)(?:usd|eur|gbp|inr|jpy|cad|aud|chf|cny)?$/.test(
    column.name,
  );

/** A bounded, conservative business-job grammar. Bindings are visible; ambiguity never guesses. */
export function analyzeReconciliationIntent(
  query: string,
  workbook: Workbook,
): ReconciliationIntent | undefined {
  if (
    !isReconciliationRequest(query) ||
    /\b(?:how\s+(?:to|do)|what\s+is|explain|syntax)\b/i.test(query)
  )
    return undefined;
  if (isComplexRequest(query)) return undefined;
  if (
    /\b(?:forecast|predict|regression|convert\s+(?:currenc|usd|eur)|foreign\s+exchange)\b/i.test(
      query,
    )
  ) {
    return question(
      'This built-in job reconciles two sheets and creates a complete exception report; it does not support currency conversion or forecasts.',
      'Start with a reconciliation-only request in a single currency.',
    );
  }
  const sheets = workbook.sheets;
  if (sheets.length < 2)
    return question(
      'A reconciliation needs two sheets in the same workbook. Your source data will not be modified.',
      'Upload a workbook containing both source tables, then ask me to reconcile them.',
    );
  const leftNamed = field(query, 'left sheet');
  const rightNamed = field(query, 'right sheet');
  const mentions = sheetMentions(query, sheets);
  const left = leftNamed
    ? sheets.find((sheet) => sheet.name.toLowerCase() === decode(leftNamed).toLowerCase())
    : (mentions[0] ?? (mentions.length === 0 && sheets.length === 2 ? sheets[0] : undefined));
  const right = rightNamed
    ? sheets.find((sheet) => sheet.name.toLowerCase() === decode(rightNamed).toLowerCase())
    : (mentions[1] ?? (mentions.length === 0 && sheets.length === 2 ? sheets[1] : undefined));
  if (!left || !right || left === right || (!leftNamed && !rightNamed && mentions.length > 2)) {
    const pairs = sheets.flatMap((first, index) =>
      sheets.slice(index + 1).map((second) => [first, second] as const),
    );
    return question(
      'Choose the two source sheets; a reconciliation report always accounts for both sides.',
      'Which two sheets should I reconcile?',
      pairs.slice(0, 8).map(([first, second]) => ({
        label: `${first.name} ↔ ${second.name}`,
        query: append(
          query,
          `left sheet: ${quote(first.name)}; right sheet: ${quote(second.name)}`,
        ),
        description: 'Read both source tables; create new report sheets only',
      })),
    );
  }
  const commonHeader = Number(last(query, /\bheaders?\s+(?:(?:on|in)\s+)?row\s*(\d+)/gi) ?? 1);
  const leftHeaderRow = Number(field(query, 'left header row') ?? commonHeader);
  const rightHeaderRow = Number(field(query, 'right header row') ?? commonHeader);
  if (
    !Number.isSafeInteger(leftHeaderRow) ||
    leftHeaderRow < 1 ||
    leftHeaderRow > left.rows.length ||
    !Number.isSafeInteger(rightHeaderRow) ||
    rightHeaderRow < 1 ||
    rightHeaderRow > right.rows.length
  )
    return question(
      'The header-row binding is outside a source sheet.',
      'Specify valid positive header rows, for example “left header row: 2; right header row: 1”.',
    );
  const leftColumns = columns(left, leftHeaderRow);
  const rightColumns = columns(right, rightHeaderRow);
  const sharedKeys = leftColumns.filter(identifier).flatMap((candidate) => {
    const matches = rightColumns.filter((column) => column.name === candidate.name);
    const leftMatches = leftColumns.filter((column) => column.name === candidate.name);
    return matches.length === 1 && leftMatches.length === 1
      ? [{ left: candidate, right: matches[0]! }]
      : [];
  });
  const explicitShared = last(
    query,
    /\b(?:by|using|on)\s+(?:keys?\s+)?(.+?)(?=;|\s+(?:compare|comparing|with\s+(?:an?\s+)?tolerance|tolerance|keys?\s+only)\b|$)/gi,
  );
  const leftKeyName = field(query, 'left keys?') ?? explicitShared;
  const rightKeyName = field(query, 'right keys?') ?? explicitShared;
  const leftKeys = leftKeyName
    ? resolveKeys(leftKeyName, leftColumns)
    : sharedKeys.length === 1
      ? [sharedKeys[0]!.left.column]
      : undefined;
  const rightKeys = rightKeyName
    ? resolveKeys(rightKeyName, rightColumns)
    : sharedKeys.length === 1
      ? [sharedKeys[0]!.right.column]
      : undefined;
  if (!leftKeys || !rightKeys || leftKeys.length !== rightKeys.length) {
    const suggested = sharedKeys.length
      ? sharedKeys
      : leftColumns
          .filter(identifier)
          .flatMap((first) =>
            rightColumns.filter(identifier).map((second) => ({ left: first, right: second })),
          );
    return question(
      'A key must identify the same business record on both sides. Repeated keys will be exceptions, not arbitrary first matches.',
      'Which columns identify the same record? You can also type separate left/right keys or composite keys.',
      suggested.map((pair) => ({
        label: `${columnLabel(pair.left)} ↔ ${columnLabel(pair.right)}`,
        query: append(query, `left keys: ${pair.left.column}; right keys: ${pair.right.column}`),
        description: 'Unique one-to-one keys only; duplicates are reported',
      })),
    );
  }
  const keysOnly =
    /\b(?:keys?\s+only|identifiers?\s+only|without\s+amounts?|no\s+amount\s+comparison)\b/i.test(
      query,
    );
  const compare = last(query, /\b(?:compare|comparing)\s+(?:amounts?\s*:\s*)?([^;]+)/gi);
  const comparison = compare?.split(/\s+(?:to|with|against|vs\.?|and)\s+/i);
  const leftAmountName =
    field(query, 'left amount') ?? (comparison?.length === 2 ? comparison[0] : undefined);
  const rightAmountName =
    field(query, 'right amount') ?? (comparison?.length === 2 ? comparison[1] : undefined);
  const leftAmounts = leftColumns.filter(
    (column) => amount(column) && !leftKeys.includes(column.column),
  );
  const rightAmounts = rightColumns.filter(
    (column) => amount(column) && !rightKeys.includes(column.column),
  );
  const leftAmount = leftAmountName
    ? resolveColumn(leftAmountName, leftColumns)
    : leftAmounts.length === 1
      ? leftAmounts[0]!.column
      : undefined;
  const rightAmount = rightAmountName
    ? resolveColumn(rightAmountName, rightColumns)
    : rightAmounts.length === 1
      ? rightAmounts[0]!.column
      : undefined;
  if (!keysOnly && (!leftAmount || !rightAmount)) {
    const pairs = leftAmounts.flatMap((first) => rightAmounts.map((second) => ({ first, second })));
    return question(
      'Choose comparable numeric amount columns, or explicitly request key-only reconciliation. I will not assume a currency conversion or treat missing values as zero.',
      'What amounts should be compared?',
      [
        ...pairs.slice(0, 7).map(({ first, second }) => ({
          label: `${columnLabel(first)} ↔ ${columnLabel(second)}`,
          query: append(query, `left amount: ${first.column}; right amount: ${second.column}`),
          description: 'Compare numeric amounts in the same unit',
        })),
        {
          label: 'Compare keys only',
          query: append(query, 'keys only'),
          description: 'Membership and duplicate checks; no amount comparison',
        },
      ],
    );
  }
  const toleranceText = last(query, /\b(?:absolute\s+)?tolerance\s*(?:of|:|=)?\s*([^;\s]+)/gi);
  const tolerance = toleranceText === undefined ? 0 : Number(toleranceText);
  if (
    !Number.isFinite(tolerance) ||
    tolerance < 0 ||
    /\b(?:within|tolerance).{0,25}(?:%|\bpercent\b)/i.test(query)
  )
    return question(
      'This job uses an absolute numeric tolerance in the same unit as the amounts, not a percentage.',
      'Specify an absolute tolerance, for example “tolerance: 0.01”, or use 0 for exact amounts.',
    );
  const keyNormalization = /\b(?:keys?\s+exact|exact\s+keys?|case[- ]sensitive)\b/i.test(query)
    ? 'exact'
    : /\b(?:ignore\s+(?:key\s+)?case|case[- ]insensitive|keys?\s+trim_casefold)\b/i.test(query)
      ? 'trim_casefold'
      : /\b(?:trim\s+(?:keys?|identifiers?)|keys?\s+trim)\b/i.test(query)
        ? 'trim'
        : 'exact';
  const args: Record<string, unknown> = {
    leftSheet: left.name,
    rightSheet: right.name,
    leftKeys,
    rightKeys,
    leftHeaderRow,
    rightHeaderRow,
    tolerance,
    keyNormalization,
    ...(!keysOnly ? { leftAmount, rightAmount } : {}),
  };
  const bindings = [
    {
      label: 'Left key columns',
      value: leftKeys
        .map((key) => columnLabel(leftColumns.find((column) => column.column === key)!))
        .join(' + '),
    },
    {
      label: 'Right key columns',
      value: rightKeys
        .map((key) => columnLabel(rightColumns.find((column) => column.column === key)!))
        .join(' + '),
    },
    {
      label: 'Amount comparison',
      value: keysOnly
        ? 'Not requested (keys only)'
        : `${left.name}!${leftAmount} ↔ ${right.name}!${rightAmount}; absolute tolerance ${tolerance}`,
    },
    {
      label: 'Key matching',
      value:
        keyNormalization === 'exact'
          ? 'Exact typed values; case and whitespace are significant'
          : keyNormalization === 'trim'
            ? 'Trim string keys; preserve case and types'
            : 'Trim and case-fold string keys; preserve types',
    },
    {
      label: 'Header rows',
      value: `${left.name}: ${leftHeaderRow}; ${right.name}: ${rightHeaderRow}`,
    },
  ];
  return {
    message: `### Reconciliation report: ${left.name} ↔ ${right.name}\n\nI will create four new sheets: **Summary**, **Matched records**, **Exceptions**, and **Methodology**. Original sheets remain unchanged.\n\n${bindings.map((binding) => `- **${binding.label}:** ${binding.value}`).join('\n')}\n\nEvery source row after the selected headers is accounted for. Duplicate or invalid keys are exceptions, never guessed matches. Missing, nonnumeric and unavailable formula amounts are disclosed rather than counted as zero. These are report bindings—not completed reconciliation results. Review the preview and apply to produce the finished workbook.`,
    proposedAction: {
      name: 'reconcile_sheets',
      category: 'transform',
      args,
      explanation: `Produce a complete, source-preserving reconciliation workbook for ${left.name} and ${right.name}.`,
    },
    evidence: [
      {
        id: 'reconciliation-bindings',
        kind: 'inspection',
        title: 'Inspected reconciliation bindings',
        source: `${left.name}!A${leftHeaderRow}:${indexToColumn(Math.max(0, leftColumns.length - 1))}${Math.max(leftHeaderRow, left.rows.length)}; ${right.name}!A${rightHeaderRow}:${indexToColumn(Math.max(0, rightColumns.length - 1))}${Math.max(rightHeaderRow, right.rows.length)}`,
        facts: bindings,
        note: 'Default header row is 1 unless explicitly selected. Amounts must have the same unit; no FX conversion. Report outputs are computed only after engine verification.',
      },
    ],
  };
}
