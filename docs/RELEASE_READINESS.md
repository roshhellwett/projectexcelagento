# Release readiness

ExcelAgento is a browser spreadsheet copilot with a deterministic cleaning and analysis
core. This document distinguishes the implemented release scope from the broader product
ambition of supporting an analyst's entire workflow.

## Implemented and regression-tested

- Workbook import/export: `.xlsx`, legacy binary `.xls`, and delimited text; multiline
  quoted fields, escaped literal quotes, UTF encodings, leading-zero identifiers, unsafe
  integer preservation, decimal precision, date formats, and 1900/1904 epochs.
- Import limit: 1,500,000 allocated cells per workbook, including rectangular CSV padding;
  UI upload limit: 50 MB.
- Cleaning, filtering, sorting, computed columns, lookup/join operations, and pivot-style
  summaries through the engine registry.
- Atomic multi-step execution with schema validation, bounded previews, invariant and patch
  checks, destructive-change confirmation, and a single undo/redo entry for a successful plan.
  Failed plans preserve existing history and redo branches.
- Descriptive statistics: missing/nonnumeric counts, mean, median, quartiles, sample variance,
  sample/population standard deviation, and IQR-based outliers with original row numbers.
- Pearson correlation and single-predictor linear regression with an intercept, R², residual
  standard error, and complete-pair exclusions. Insufficient or constant data is explicit.
- Live evaluation of supported formulas for statistics and engine aggregates; formula errors
  are excluded and counted, rather than silently converted to numbers.
- Local learned-action persistence with verified success/failure counts. Cloud memory is
  disabled unless the deployment supplies both Supabase settings.
- Browser-local IndexedDB checkpoints with restore/discard, transaction-commit status, retry,
  and explicit storage-failure warnings. Checkpoints preserve Dates and the workbook epoch;
  restoration starts fresh undo history and conversation, not a recovered agent session.
- Workbook generation/revision fences: replacement clears proposals and conversation; edits
  invalidate outstanding previews and confirmation. Late callbacks from previous turns are ignored.
- Mandatory strict critic approval, one bounded repair-and-review attempt, and read-only inspection
  tools for specialists. Incomplete provider answers never receive fabricated financial conclusions.
- Native Gemini tool calls/results, caller cancellation across retry/tool phases, and complete
  provider-reported usage totals for successful inference calls across a turn. Missing usage is unknown.
- SheetJS CE 0.20.3 pinned from its authoritative distribution with lockfile integrity.
- 1904 metadata preserved through edit/undo/redo and patch verification; current supported formula
  caches and actual full-recalculation XML on export. Unsupported/error caches are omitted.
- Full-operation tall-sheet filter/extraction/deduplication regressions at 150,000 rows.

## Verification commands

From the repository root, with Node 22+ and the pinned pnpm version:

```bash
pnpm install --frozen-lockfile
pnpm lint
pnpm typecheck
pnpm test
pnpm evals
pnpm build
```

CI runs these checks. Tests include real workbook round trips, property-based undo/redo,
recorded provider contracts, and jsdom workspace integration. Statistical tests include
known reference distributions, perfect and imperfect regressions, missing pairs, constant
columns, large offsets, tiny values, numerical overflow, and live formula dependencies.

## Try the analyst workflows

Choose **Try the local agent**, upload a tabular workbook, then ask:

```text
descriptive statistics for column D
find outliers in column D
correlation between column C and column D
linear regression of column D on column C
```

The local chat assumes the first row is the header. Connected models can supply an explicit
`headerRow` to the statistical read tools. Regression models D as the response and C as the
predictor. Outlier detection reports observations without deleting them.

## Remaining release gaps

1. **Real-browser coverage:** the repository has jsdom integration tests, but no Playwright
   suite validating real-browser workers, clipboard behavior, downloads, and rendering.
2. **Excel fidelity:** the workbook model does not preserve all Excel features. Merged cells,
   comments, hyperlinks, layout, charts, styles, macros, named ranges, and conditional
   formatting need dedicated preservation/reporting work. Import reports some dropped
   features; it is not a complete fidelity inventory. Formula-bearing export that requires sheet
   renaming is now rejected explicitly. Typed error-cell fidelity remains incomplete.
3. **Formula semantics:** the evaluator implements a subset of Excel. Structural operations
   on formula-bearing workbooks are now blocked conservatively, including formulas on other
   sheets. This protects meaning but restricts workflows until an Excel-aware reference rewriter
   exists. Unsupported formulas need Excel/LibreOffice recalculation; application behavior still
   requires verification. This release does not claim complete Excel compatibility.
4. **Large-workbook responsiveness:** imports have limits and parsing/export use a worker,
   but operation validation, repeated workbook cloning, and several analyses still run on
   the main thread. Benchmarking and worker-based execution remain important for large files.
5. **Recovery scope:** one latest successful checkpoint is available per browser/origin.
   Browser storage can be cleared or fail, and a crash before checkpoint commit can lose recent
   changes. Undo history and chat are intentionally not persisted. Export remains essential;
   multi-workbook libraries and conflict-safe multi-tab checkpoints remain product work.
6. **Optional integrations:** live model-provider compatibility and Supabase provisioning,
   isolation, and access policies have not been verified against deployed services. No API
   key or backend is needed for the tested local workflows.
7. **Broader data science:** multivariate models, hypothesis testing, forecasting, model
   training/validation, notebook execution, and chart/dashboard authoring remain product work.

The production bundle can be deployed to Vercel for a scoped pilot. Public general availability
should follow real-browser validation, deployment-specific provider/Supabase verification, and
workflows that match the supported fidelity boundaries above. Changes are local until a separate
release decision; this document is not a claim that the deployed production app has been upgraded.
