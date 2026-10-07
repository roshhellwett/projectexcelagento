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
- Local learned-action persistence with verified success/failure counts. All memory and
  working-session scratchpads operate 100% locally in-browser with zero database dependencies.
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
- Mission control (`#/missions`): latest 50 browser-local requests, plans/actions, answers, evidence
  and execution receipts. Planning/executing/prepared/analyzed/applied/undone/failed/stale/cancelled/
  interrupted states, committed save/delete/clear status, in-memory fallback, retry and JSON download.
- Prepared refresh resume requires full-content SHA-256 matching, fresh schema/engine previews,
  generation/revision fencing and renewed confirmation. Closed-tab active tasks become interrupted;
  completed/cancelled/stale work cannot execute again. Replanning starts a new reviewable request.
- Dedicated workers stage chat mutations where supported, emit per-step verification progress and
  terminate on abort/crash/timeout; atomic UI commit happens only for the unchanged live document.
- Integrated Excel-style workspace ribbon with Home, Insert, Draw, Page Layout, Formulas, Data,
  Review, Automate, and Help surfaces. Actionable ribbon requests route through the existing
  agent review and engine guardrails; unsupported drawing, connection, and rich-fidelity controls
  remain visible but disabled with an explanation instead of silently pretending to work.
- Manual operation access exposes the live engine catalog and examples, including operations beyond
  the guided form through validated JSON arguments. This keeps the workspace and model-facing
  operation set aligned as the registry grows.
- Task-specific receipts track actual history undo/redo/reset. Discarded redo branches and compacted
  operations cannot accidentally undo another task. Undo bindings are not persisted across refresh.
- Deterministic analyst briefing: source-backed full-column distributions/quality/exclusions,
  bounded categorical/histogram charts with table alternatives, unsupported/error/volatile formula
  exclusion, and exact positional baseline/shape/epoch comparison. Baseline is memory-only.
- Real Chromium mission reload/resume/confirmation/worker/undo/redo/delete and stale-edit flows,
  plus responsive mobile briefing rendering. Strict Mode cell-editor lifecycle regression fixed.

## Verification commands

From the repository root, with Node 22+ and the pinned pnpm version:

```bash
pnpm install --frozen-lockfile
pnpm lint
pnpm typecheck
pnpm test
pnpm evals
pnpm build
pnpm exec playwright install --with-deps chromium
pnpm test:e2e
```

CI is configured to run these checks. Tests include real workbook round trips, property-based undo/redo,
recorded provider contracts, fake-IndexedDB transaction read-back/failure tests, jsdom workspace
integration, and real Chromium browser flows. The Playwright runner builds and serves the production bundle through Vite preview;
this is not a deployed production/CSP or cross-browser verification. Statistical tests include
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

1. **Remaining browser coverage:** Chromium now covers mission persistence/recovery, real workers,
   confirmation/undo/redo, deletion, stale edits and mobile briefing rendering. Firefox/WebKit,
   clipboard, downloaded XLSX validation, large-file responsiveness and deployed CSP remain unchecked.
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
4. **Large-workbook responsiveness:** parsing/export and chat mutations use workers where available,
   but proposal guardrails, secure-hash serialization, manual grid operations, briefing analysis,
   repeated cloning and history commit still involve main-thread work. No heavy-use latency/memory
   benchmark has been established; the worker fallback cannot stop synchronous CPU work.
5. **Recovery scope:** one latest successful checkpoint and up to 50 mission records are available
   per browser/origin. They commit independently, so a crash may leave a stale in-progress mission
   status or an older workbook checkpoint. Browser storage can fail or be cleared; no active call
   resumes automatically. Chat and undo history are not restored; receipts are not persisted undo
   handles. Records are unencrypted and may contain customer data. Export remains essential;
   multi-workbook libraries, crash-consistent journaling and conflict-safe multi-tab state remain work.
6. **Optional integrations:** live model-provider compatibility is optional via BYOK.
   No API key or backend database is needed for the tested local workflows.
7. **Broader data science:** multivariate models, hypothesis testing, forecasting, model
   training/validation, notebook execution, and chart/dashboard authoring remain product work.
8. **Official licensing deployment:** the Supabase licensing schema and Edge Function are in the
   repository, but production access depends on applying both licensing migrations, setting the
   server-only peppers and `APP_ORIGIN`, deploying the function, and bootstrapping the first admin.
   The static SPA cannot prove that the remote Supabase project has those settings; verify them with a
   real production sign-in and an admin key-generation/activation smoke test.

The production bundle can be deployed to Vercel for a scoped pilot. Public general availability
should follow real-browser validation, deployment-specific provider verification, and
workflows that match the supported fidelity boundaries above. Changes are local until a separate
release decision; this document is not a claim that the deployed production app has been upgraded.
