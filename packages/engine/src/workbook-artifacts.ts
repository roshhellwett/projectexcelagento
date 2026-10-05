import type { WorkbookArtifact } from './types.js';
import { columnToIndex } from './workbook.js';
const object = (value: unknown): value is Record<string, unknown> => !!value && typeof value === 'object' && !Array.isArray(value);
const text = (value: unknown, limit = 4000): value is string => typeof value === 'string' && value.length <= limit;
const row = (value: unknown): value is number => typeof value === 'number' && Number.isInteger(value) && value >= 1 && value <= 1048576;
const column = (value: unknown): value is string => typeof value === 'string' && /^[A-Z]+$/.test(value) && (columnToIndex(value) ?? 16384) < 16384;

/** Metadata is a historical report, not authorization to open, execute or export a workbook. */
export function validateWorkbookArtifacts(value: unknown): WorkbookArtifact[] {
  if (!Array.isArray(value) || value.length > 25) throw new Error('Saved workbook deliverables are invalid.');
  return value.map((item): WorkbookArtifact => {
    if (!object(item) || !text(item.id, 1000) || !text(item.title, 1000) || item.kind !== 'reconciliation' ||
      !Array.isArray(item.sheets) || item.sheets.length < 1 || item.sheets.length > 10 || !item.sheets.every((sheet) => object(sheet) && text(sheet.name, 31) && !!sheet.name && ['summary','matched','exceptions','methodology'].includes(String(sheet.role))) ||
      !Array.isArray(item.sources) || item.sources.length > 20 || !item.sources.every((source) => object(source) && text(source.sheet, 1000) && row(source.startRow) && row(source.endRow) && source.endRow >= source.startRow && column(source.startColumn) && column(source.endColumn) && columnToIndex(source.endColumn)! >= columnToIndex(source.startColumn)!) ||
      !Array.isArray(item.facts) || item.facts.length > 50 || !item.facts.every((fact) => object(fact) && text(fact.label) && text(fact.value, 10000)) ||
      !Array.isArray(item.checks) || item.checks.length > 25 || !item.checks.every((check) => object(check) && text(check.id, 1000) && text(check.label) && ['passed','warning','failed'].includes(String(check.status)) && text(check.detail, 10000)) ||
      !Array.isArray(item.notes) || item.notes.length > 30 || !item.notes.every((note) => text(note, 10000))) throw new Error('Saved workbook deliverable metadata is invalid.');
    return { id: item.id, title: item.title, kind: 'reconciliation',
      sheets: item.sheets.map((sheet) => ({ name: sheet.name, role: sheet.role })),
      sources: item.sources.map((source) => ({ sheet: source.sheet, startRow: source.startRow, endRow: source.endRow, startColumn: source.startColumn, endColumn: source.endColumn })),
      facts: item.facts.map((fact) => ({ label: fact.label, value: fact.value })),
      checks: item.checks.map((check) => ({ id: check.id, label: check.label, status: check.status, detail: check.detail })), notes: [...item.notes],
    };
  });
}
