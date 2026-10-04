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

function cellCount(workbook: Workbook): number {
  let total = 0;
  for (const sheet of workbook.sheets) {
    for (const row of sheet.rows) {
      total += row.length;
    }
  }
  return total;
}

/**
 * Approximate retained size of a patch in cells. A cell patch keeps two values,
 * a structural patch keeps two whole workbooks. Measuring cells instead of bytes
 * keeps the budget allocation-free and cheap to track incrementally.
 */
function patchCost(patch: Patch): number {
  let total = 0;
  for (const entry of patch) {
    total +=
      entry.kind === 'workbook' ? cellCount(entry.oldWorkbook) + cellCount(entry.newWorkbook) : 1;
  }
  return total;
}

/**
 * A retained history slot. `collapsed` is a full-workbook snapshot that stands in
 * for the operations that were compacted away, which keeps undo working across
 * the compaction boundary instead of losing that part of the history.
 */
type RetainedEntry =
  | { kind: 'patch'; operationName: string; patch: Patch; cost: number }
  | { kind: 'collapsed'; operationName: string; workbook: Workbook; cost: number };

const COMPACTED_OPERATION_NAME = 'history.compacted';

export interface HistoryOptions {
  /** Retain a full snapshot every N operations (used to speed up replay). */
  snapshotEvery?: number;
  /** Maximum number of retained operations before the oldest ones are compacted. */
  maxEntries?: number;
  /**
   * Approximate retained-cell budget. Exceeding it compacts the oldest retained
   * operations into a single snapshot. Defaults to 1,000,000 cells.
   */
  maxRetainedCells?: number;
}

const DEFAULT_MAX_ENTRIES = 200;
const DEFAULT_MAX_RETAINED_CELLS = 1_000_000;

function positiveInteger(value: number | undefined, fallback: number, label: string): number {
  const resolved = value ?? fallback;
  if (!Number.isInteger(resolved) || resolved < 1) {
    throw new Error(`${label} must be a positive integer.`);
  }
  return resolved;
}

/**
 * A linear operation history. The cursor points between entries: 0 is the
 * initial workbook and entries.length is the latest workbook.
 *
 * Retention is bounded. Once the entry count or the approximate cell budget is
 * exceeded, the oldest retained operations are replaced by a single snapshot
 * entry, so memory stops growing while undo still crosses the boundary between
 * the compacted block and the operations kept after it.
 */
export class HistoryStack {
  private entries: RetainedEntry[] = [];

  private cursor = 0;

  private readonly snapshots = new Map<number, Workbook>();

  private readonly snapshotEvery: number;

  private readonly maxEntries: number;

  private readonly maxRetainedCells: number;

  private retainedCells = 0;

  private compactedOperations = 0;

  constructor(initial: Workbook, options: HistoryOptions = {}) {
    const snapshotEvery = positiveInteger(options.snapshotEvery, 10, 'snapshotEvery');
    this.snapshotEvery = snapshotEvery;
    this.maxEntries = positiveInteger(options.maxEntries, DEFAULT_MAX_ENTRIES, 'maxEntries');
    this.maxRetainedCells = options.maxRetainedCells ?? DEFAULT_MAX_RETAINED_CELLS;
    if (!Number.isFinite(this.maxRetainedCells) || this.maxRetainedCells < 1) {
      throw new Error('maxRetainedCells must be a positive number.');
    }
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

  /** Operations that were replaced by snapshots to stay inside the budget. */
  get compactedCount(): number {
    return this.compactedOperations;
  }

  get history(): HistoryEntry[] {
    return this.entries.map((entry) => ({
      operationName: entry.operationName,
      patch: entry.kind === 'patch' ? clonePatch(entry.patch) : [],
    }));
  }

  /** Record an already-applied operation and discard any redo branch. */
  commit(operationName: string, after: Workbook, patch: Patch, _inverse?: Patch): Workbook {
    if (this.cursor < this.entries.length) {
      this.entries = this.entries.slice(0, this.cursor);
      for (const key of this.snapshots.keys()) {
        if (key > this.cursor) {
          this.snapshots.delete(key);
        }
      }
      if (this.entries[0]?.kind !== 'collapsed') {
        this.compactedOperations = 0;
      }
      this.recomputeCost();
    }

    const cost = patchCost(patch);
    this.entries.push({
      kind: 'patch',
      operationName,
      patch: clonePatch(patch),
      cost,
    });
    this.cursor += 1;
    this.retainedCells += cost;
    if (this.cursor % this.snapshotEvery === 0) {
      this.snapshots.set(this.cursor, cloneWorkbook(after));
    }
    this.enforceBudget();
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
    // Position 0 is always available as a snapshot: the constructor records the
    // origin, and compaction replaces it with the compacted state.
    let nearestSnapshot = 0;
    for (const key of this.snapshots.keys()) {
      if (key <= position && key >= nearestSnapshot) {
        nearestSnapshot = key;
      }
    }
    let workbook = cloneWorkbook(this.snapshots.get(nearestSnapshot) as Workbook);
    for (let index = nearestSnapshot + 1; index <= position; index += 1) {
      const entry = this.entries[index - 1];
      if (entry === undefined) {
        continue;
      }
      workbook =
        entry.kind === 'collapsed'
          ? cloneWorkbook(entry.workbook)
          : applyPatch(workbook, entry.patch);
    }
    return workbook;
  }

  private recomputeCost(): void {
    let total = 0;
    for (const entry of this.entries) {
      total += entry.cost;
    }
    this.retainedCells = total;
  }

  /**
   * Compact the oldest retained operations into snapshots until the history is
   * back inside both budgets.
   */
  private enforceBudget(): void {
    while (this.overBudget()) {
      const block = this.compactionBlock();
      if (block === 0) {
        return;
      }
      this.compact(block);
    }
  }

  private overBudget(): boolean {
    if (this.entries.length === 0) {
      return false;
    }
    if (this.entries.length > this.maxEntries) {
      return true;
    }
    if (this.retainedCells <= this.maxRetainedCells) {
      return false;
    }
    // A single collapsed snapshot is the floor: it is the cheapest complete
    // representation of the workbook, so no further compaction can help.
    return this.entries.length > 1;
  }

  /**
   * How many of the oldest retained operations to replace with one snapshot: at
   * least one, and enough to bring the entry count and the retained-cell
   * estimate back under their limits. Compacting a prefix of `block` operations
   * leaves `entries.length - block + 1` retained entries.
   */
  private compactionBlock(): number {
    const count = this.entries.length;
    if (count === 0) {
      return 0;
    }
    // Re-compacting a snapshot into itself saves nothing, so when the oldest entry
    // is already a snapshot the block has to reach the next entry as well.
    let block = this.entries[0]?.kind === 'collapsed' ? Math.min(2, count) : 1;
    let remainingCells = this.retainedCells;
    for (let index = 0; index < block; index += 1) {
      remainingCells -= this.entries[index]?.cost ?? 0;
    }
    while (
      block < count &&
      (count - block + 1 > this.maxEntries || remainingCells > this.maxRetainedCells)
    ) {
      block += 1;
      remainingCells -= this.entries[block - 1]?.cost ?? 0;
    }
    return block;
  }

  /** Replace the oldest `block` retained operations with one snapshot entry. */
  private compact(block: number): void {
    const tail = this.entries[block - 1];
    const workbook =
      tail !== undefined && tail.kind === 'collapsed' ? tail.workbook : this.materialize(block);
    this.entries = [
      {
        kind: 'collapsed',
        operationName: COMPACTED_OPERATION_NAME,
        workbook,
        cost: cellCount(workbook),
      },
      ...this.entries.slice(block),
    ];
    for (const key of [...this.snapshots.keys()]) {
      if (key < block) {
        this.snapshots.delete(key);
        continue;
      }
      // Compacting `block` entries removes `block - 1` positions from the retained
      // window, so an old position `k` becomes `k - block + 1` - otherwise a later
      // undo would restore a workbook from the wrong operation.
      const snapshot = this.snapshots.get(key);
      this.snapshots.delete(key);
      if (snapshot !== undefined) {
        this.snapshots.set(key - block + 1, snapshot);
      }
    }
    // Position 0 is the base every replay starts from, so it becomes the
    // compacted state. The collapsed entry and this snapshot share one workbook:
    // both are only ever read through `cloneWorkbook`.
    this.snapshots.set(0, workbook);
    this.compactedOperations += block - 1;
    // commit() always leaves the cursor on the newest operation, so the compacted
    // block sits strictly behind it.
    this.cursor = this.entries.length;
    this.recomputeCost();
  }
}
