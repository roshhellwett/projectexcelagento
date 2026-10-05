import type { CompactColumnProfile } from '@excel-agent/agent';

export interface WorkspaceWorkflow {
  id: string;
  title: string;
  description: string;
  category: 'Clean' | 'Analyze' | 'Organize';
  prompt: string;
  icon:
    | 'sparkles'
    | 'duplicates'
    | 'missing'
    | 'statistics'
    | 'outliers'
    | 'sort'
    | 'dates'
    | 'overview';
}

/** Real, sheet-aware entry points into the existing planner. No fixed sample column names. */
export function workspaceWorkflows(profiles: CompactColumnProfile[]): WorkspaceWorkflow[] {
  const numeric = profiles.find((column) => column.isNumeric);
  const text = profiles.find(
    (column) => !column.isNumeric && !column.isDate && column.nonBlankCount > 0,
  );
  const date = profiles.find((column) => column.isDate || /date|time/i.test(column.rawName));
  const sortColumn = numeric ?? profiles[0];
  const workflows: WorkspaceWorkflow[] = [
    {
      id: 'duplicates',
      title: 'Remove duplicates',
      description: 'Find repeated records. Review the rows before removing them.',
      category: 'Clean',
      prompt: 'remove duplicate rows',
      icon: 'duplicates',
    },
    {
      id: 'missing',
      title: 'Find missing values',
      description: 'See which columns need attention before your next analysis.',
      category: 'Clean',
      prompt: 'audit missing values and empty cells',
      icon: 'missing',
    },
    {
      id: 'overview',
      title: 'Understand this sheet',
      description: 'Get a grounded overview of your columns and records.',
      category: 'Analyze',
      prompt: 'give me an overview of this worksheet',
      icon: 'overview',
    },
    {
      id: 'statistics',
      title: 'Describe the numbers',
      description: numeric
        ? `Mean, median, spread, and quartiles for ${numeric.rawName}.`
        : 'Choose a column to inspect its statistical distribution.',
      category: 'Analyze',
      prompt: numeric
        ? `descriptive statistics for column ${numeric.letter}`
        : 'descriptive statistics',
      icon: 'statistics',
    },
    {
      id: 'outliers',
      title: 'Spot unusual values',
      description: 'Flag IQR outliers with original row numbers. Keep your data intact.',
      category: 'Analyze',
      prompt: numeric ? `find outliers in column ${numeric.letter}` : 'find outliers',
      icon: 'outliers',
    },
  ];
  if (text)
    workflows.push({
      id: 'text',
      title: 'Tidy text',
      description: `Remove extra whitespace in ${text.rawName}.`,
      category: 'Clean',
      prompt: `trim whitespace in column ${text.letter}`,
      icon: 'sparkles',
    });
  if (sortColumn)
    workflows.push({
      id: 'sort',
      title: 'Bring order to your data',
      description: `Sort the entire dataset by ${sortColumn.rawName}.`,
      category: 'Organize',
      prompt: `sort by column ${sortColumn.letter} descending`,
      icon: 'sort',
    });
  if (date)
    workflows.push({
      id: 'dates',
      title: 'Standardize dates',
      description: `Normalize ${date.rawName} to a consistent ISO date format.`,
      category: 'Organize',
      prompt: `format dates in column ${date.letter} to YYYY-MM-DD`,
      icon: 'dates',
    });
  return workflows;
}
