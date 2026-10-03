import type { HistoryEntry, Patch, PatchEntry, Workbook } from './types.js';
import { applyPatch, cloneWorkbook } from './workbook.js';

function clonePatchEntry(entry: PatchEntry): PatchEntry {
  if (entry.kind === 'workbook') {
    return {
      kind: 'workbook',
      oldWorkbook: cloneWorkbook(entry.oldWorkbook),
      newWorkbook: cloneWorkbook(entry.newWorkbook),
    };
  }
  const cloneValue = (v: (typeof entry)['oldValue']) =>
    v instanceof Date ? new Date(v.getTime()) : v;
  return {
    ...entry,
    address: { ...entry.address },
    oldValue: cloneValue(entry.oldValue),
    newValue: cloneValue(entry.newValue),
  };
}

function clonePatch(patch: Patch): Patch {
  return patch.map(clonePatchEntry);
}

export interface HistoryOptions {
  snapshotEvery?: number;
}

/**
 * A linear operation history. The cursor points between entries: 0 is the
 * initial workbook and entries.length is the latest workbook.
 */
export class HistoryStack {
  private entries: HistoryEntry[] = [];

  private cursor = 0;

  private readonly snapshots = new Map<number, Workbook>();

  private readonly snapshotEvery: number;

  constructor(initial: Workbook, options: HistoryOptions = {}) {
    const snapshotEvery = options.snapshotEvery ?? 10;
    if (!Number.isInteger(snapshotEvery) || snapshotEvery < 1) {
      throw new Error('snapshotEvery must be a positive integer.');
    }
    this.snapshotEvery = snapshotEvery;
    this.snapshots.set(0, cloneWorkbook(initial));
  }

  get position(): number {
    return this.cursor;
  }

  get length(): number {
    return this.entries.length;
  }

  get canUndo(): boolean {
    return this.cursor > 0;
  }

  get canRedo(): boolean {
    return this.cursor < this.entries.length;
  }

  get snapshotCount(): number {
    return this.snapshots.size;
  }

  get history(): HistoryEntry[] {
    return this.entries.map((entry) => ({
      operationName: entry.operationName,
      patch: clonePatch(entry.patch),
      inverse: clonePatch(entry.inverse),
    }));
  }

  /** Record an already-applied operation and discard any redo branch. */
  commit(operationName: string, after: Workbook, patch: Patch, inverse: Patch): Workbook {
    if (this.cursor < this.entries.length) {
      this.entries = this.entries.slice(0, this.cursor);
      for (const key of this.snapshots.keys()) {
        if (key > this.cursor) {
          this.snapshots.delete(key);
        }
      }
    }

    this.entries.push({ operationName, patch: clonePatch(patch), inverse: clonePatch(inverse) });
    this.cursor += 1;
    if (this.cursor % this.snapshotEvery === 0) {
      this.snapshots.set(this.cursor, cloneWorkbook(after));
    }
    return cloneWorkbook(after);
  }

  undo(): Workbook | undefined {
    if (!this.canUndo) {
      return undefined;
    }
    this.cursor -= 1;
    return this.materialize(this.cursor);
  }

  redo(): Workbook | undefined {
    if (!this.canRedo) {
      return undefined;
    }
    this.cursor += 1;
    return this.materialize(this.cursor);
  }

  /** Restore the workbook at an operation boundary without deleting history. */
  stepBack(position: number): Workbook {
    if (!Number.isInteger(position) || position < 0 || position > this.entries.length) {
      throw new Error(`History position must be between 0 and ${this.entries.length}.`);
    }
    this.cursor = position;
    return this.materialize(position);
  }

  restore(position: number): Workbook {
    return this.stepBack(position);
  }

  private materialize(position: number): Workbook {
    let nearestSnapshot = 0;
    for (const key of this.snapshots.keys()) {
      if (key <= position && key >= nearestSnapshot) {
        nearestSnapshot = key;
      }
    }

    let workbook = cloneWorkbook(this.snapshots.get(nearestSnapshot) as Workbook);
    for (let index = nearestSnapshot; index < position; index += 1) {
      const entry = this.entries[index];
      if (entry) {
        workbook = applyPatch(workbook, entry.patch);
      }
    }
    return workbook;
  }
}
