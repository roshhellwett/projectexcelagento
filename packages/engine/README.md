# Engine

The engine is UI-, network-, and file-system-free TypeScript. Operations receive a workbook model, validate typed arguments, produce a bounded preview, apply immutable changes, return a reversible patch, and expose invariant checks.

## Workbook and history

`Cell` values carry a value, inferred `type`, optional formula, and optional number format. Cell patches explicitly store an address plus old/new value, formula, type, and number format. Structural changes use an immutable workbook snapshot patch so row and column changes undo exactly.

`HistoryStack` stores operation patches, supports `undo`, `redo`, and `stepBack(position)`, and records a workbook snapshot every configurable number of operations (`snapshotEvery`, default `10`).

## Operation registry

`createOperationRegistry()` registers every operation by name and exposes each Zod schema. `applyOperation(workbook, name, input)` is the transactional entry point: schema validation, operation validation, preview, apply, forward/inverse patch verification, and shared invariants. An invariant failure returns a structured error and the original workbook without committing history.

## Operations

M2 includes:

- `format_dates`
- `sort_range`
- `filter_rows`
- `find_replace`
- `delete_duplicates`
- `rename_column`
- `delete_column`
- `add_column`
- `normalize_text`
- `set_cells`

The shared invariant framework checks row counts, sorted-row multisets, formula preservation, and changes outside declared target ranges. Operations that explicitly remove rows, columns, or formulas opt into the relevant exception.

## `format_dates`

```ts
const args = formatDatesArgsSchema.parse({
  sheet: 'Orders',
  column: 'B',
  format: 'YYYY-MM-DD',
});

const preview = formatDatesOperation.preview(workbook, args);
const result = formatDatesOperation.apply(workbook, args);
```

The operation understands Excel 1900/1904 serials, ISO date text, and unambiguous day-first or month-first text. Values such as `01/02/2020` are reported as `ambiguous-date`, left unchanged, and require confirmation; the operation never guesses a locale. Empty cells, non-date text, formulas, and unsupported numeric values are not overwritten. Converted cells store a UTC `Date` plus the requested `numberFormat`, which keeps the value date-like for a future Excel writer.
