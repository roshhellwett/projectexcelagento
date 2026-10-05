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
   features; it is not a complete fidelity inventory. Sheet-name sanitization on export
   also needs formula-reference rewriting when names change.
3. **Formula semantics:** the evaluator implements a subset of Excel. Structural edits need
   comprehensive reference-rewriting behavior, and unsupported formulas need a fuller
   compatibility strategy. This release does not claim complete Excel compatibility.
4. **Large-workbook responsiveness:** imports have limits and parsing/export use a worker,
   but operation validation, repeated workbook cloning, and several analyses still run on
   the main thread. Benchmarking and worker-based execution remain important for large files.
5. **Session recovery:** edited workbooks are held in memory; automatic workbook recovery
   after refresh/tab closure is not implemented. Users must export to retain their edits.
6. **Optional integrations:** live model-provider compatibility and Supabase provisioning,
   isolation, and access policies have not been verified against deployed services. No API
   key or backend is needed for the tested local workflows.
7. **Broader data science:** multivariate models, hypothesis testing, forecasting, model
   training/validation, notebook execution, and chart/dashboard authoring remain product work.

The production bundle can be deployed to Vercel for a scoped pilot. Public general
availability should follow real-browser validation and resolution of the data-fidelity,
formula-reference, and recovery gaps relevant to the intended customer workflows.
