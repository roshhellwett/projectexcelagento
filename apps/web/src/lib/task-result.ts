import type { Report } from '@excel-agent/engine';

/** Describe committed engine facts rather than repeating a model's proposed outcome. */
export function describeTaskResult(operation: string, report: Report): string {
  const quantity = (count: number, noun: string) =>
    `${count.toLocaleString()} ${noun}${count === 1 ? '' : 's'}`;
  if (report.removedRows !== undefined) {
    return `Removed ${quantity(report.removedRows, operation === 'delete_duplicates' ? 'duplicate row' : 'row')}.`;
  }
  if (report.addedRows !== undefined) return `Added ${quantity(report.addedRows, 'row')}.`;
  if (report.deletedColumns !== undefined)
    return `Removed ${quantity(report.deletedColumns, 'column')}.`;
  if (report.addedColumns !== undefined) return `Added ${quantity(report.addedColumns, 'column')}.`;
  if (report.affectedCells === 0) return 'No cells changed.';
  return `Updated ${quantity(report.affectedCells, 'cell')}.`;
}
