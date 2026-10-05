# Architecture

ExcelAgento is a pnpm monorepo with one deployable app and three libraries. The
dependency direction is strictly one-way:

```
apps/web  ->  packages/agent  ->  packages/engine
apps/web  ->  packages/engine
packages/evals -> packages/agent -> packages/engine
```

`packages/engine` knows nothing about UI or network. `packages/agent` knows nothing about
React. This keeps the correctness core testable and makes the agent reusable from a CLI,
a queue worker, or a future server runtime.

## Data flow of a single chat turn

```
User prompt
   |
   v
apps/web/src/App.tsx  (handleSendMessage)
   |
   v
apps/web/src/lib/llm-service.ts        thin adapter
   |
   v
packages/agent  ExcelAgentOrchestrator.decide()
   |   1. intent      greeting detection
   |   2. memory      replay a previously *verified* action
   |   3. heuristic   deterministic planner (analyzeSpreadsheetIntentAndData)
   |   4. llm         BYOK provider call (retry + timeout + backoff)
   |   5. guardrail   schema -> engine.validate -> preview
   v
AgentDecision { message, action?, guardrail?, evidence?, source, trace }
   |
   v
App shows a preview card; user clicks Apply
   |
   v
engine.applyOperation()  validate -> preview -> apply -> invariant check -> patch verify
   |
   v
orchestrator.learn()  reinforces or decays the association, then persists to localStorage
```

### Evidence and task receipts

Read tools and deterministic statistical paths can emit optional `EvidenceItem` records. Each record
is normalized from typed workbook output, carries a human-readable source such as `Sales!B2:B101`,
and lists exact facts and limitations. Model prose is never treated as evidence. The web workspace
renders these records in an evidence card beside the answer.

Applied operations produce a task receipt in the chat with the request, workbook generation and
revision, operation names, affected-cell counts, warnings, and status. Undo changes the receipt to
`undone`; failed operations produce a failed receipt. This separates proposed `Preview` data from
completed engine `Report` data and gives users a durable, auditable outcome for each mission.

### Why the guardrail matters

The model is never trusted. `guardrail()` runs the proposed action through three
independent gates before the UI can even offer an "Apply" button:

1. `operation.schema.safeParse` - typed contract (Zod, derived straight from the engine).
2. `operation.validate` - domain rules (sheet exists, column exists, header in range).
3. `operation.preview` - the concrete, bounded set of cell changes.

If any gate fails, the layer records the reason in `decision.trace` and the orchestrator
falls through to the next candidate or answers conversationally. An LLM hallucination
therefore degrades into a helpful message instead of a corrupted spreadsheet.

### Self-learning without magic

There is no model fine-tuning. "Self-learning" here is an honest, auditable association
store (`packages/agent/src/memory.ts`):

- A prompt is normalized (lowercased, punctuation stripped, stop words removed).
- Only associations that **actually succeeded** are recorded.
- Retrieval is Jaccard token similarity multiplied by a reliability score
  (`successes / total`, damped by observation count), so a never-proven record is never
  replayed.
- Everything round-trips through JSON, so the browser persists it in `localStorage` and
  the user can inspect or forget it from Settings.

## Invariants

`packages/engine/src/invariants.ts` provides the safety net every operation opts into:

- `rowCountUnchanged` - structural operations must declare their intent to change shape.
- `rowsMultisetEqual` - sorting must be a permutation, not a rewrite.
- `formulaCellsRemainFormulas` - formulas are never silently replaced by literals.
- `invariantNoCellsOutsideTargetRange` - no operation may touch a cell it did not declare.

`applyOperation` also verifies that the forward patch reproduces the result and that the
inverse patch restores the original. Any failure rolls back and returns a structured error
without committing to history.

## Atomic plans and statistical analysis

`engine.applyOperationPlan()` stages each operation against the previous step's result,
using the same validation, preview, invariants, and patch verification as a single operation.
Confirmation authorizes the final commit, not the private simulation. Failed and unconfirmed
plans never modify history, including an existing redo branch. A successful plan is recorded
as one change, so one undo restores its starting workbook.

`packages/engine/src/statistics.ts` provides descriptive statistics and ordinary least
squares regression. Variance and covariance use streaming centered updates; descriptive
statistics use compensated summation and inclusive interpolated quartiles. Missing and
nonnumeric values have separate counts. Insufficient observations and constant predictors
produce undefined statistics rather than invented zeros.

`createWorkbookValueReader()` memoizes live formula results for an immutable workbook
snapshot and bounds dependency recursion and range allocation. Statistical read tools and
engine aggregates use this reader. `packages/agent/src/statistical-intent.ts` exposes these
calculations through no-key chat; `statistical-tools.ts` exposes the same calculations to
connected models as `describe_column` and `analyze_column_relationship`.

## Extension points

| Goal                           | Where to change                                                                                                                   |
| ------------------------------ | --------------------------------------------------------------------------------------------------------------------------------- |
| Add an Excel operation         | `packages/engine/src/operations.ts`, then register it and document it in `packages/agent/src/tools.ts` (build fails until you do) |
| Support a new model provider   | `packages/agent/src/providers.ts` - implement `ProviderAdapter`, add to `ADAPTERS`                                                |
| Teach the agent a new phrasing | `packages/agent/src/analysis.ts`, then add an eval case in `packages/evals/src/cases.ts`                                          |
| Change the UI shell            | `apps/web/src/components/*` - components are presentational and receive props                                                     |

## Testing strategy

Five layers, each with a different job:

- **Engine** - unit tests per operation plus `fast-check` property tests for undo/redo.
- **Agent** - planner behaviour, guardrail rejection of hallucinated and engine-invalid actions,
  memory gating, and the provider adapters driven by **recorded API payloads** (Groq, OpenRouter,
  Gemini) covering 401/429/5xx, retry backoff, caller abort, timeout, and malformed bodies.
- **Evals** - a golden set of natural-language prompts mapped to expected operations, asserting
  both accuracy and that 100% of proposed actions clear the guardrail.
- **Web (unit)** - settings/usage stores (including corrupt-storage tolerance) and workbook
  import/export round-trip fidelity (values, formulas, blanks, number formats, sheet-name
  sanitization, CSV detection).
- **Web (UI)** - jsdom + Testing Library renders the real app: upload an actual `.xlsx`, export an
  actual blob, apply and undo an operation, save/clear a BYOK key, and read the usage ledger.

The UI and provider layers use a real DOM and real payload shapes rather than mocks of our own
code, so a regression in wiring is caught. CI needs no API key and makes no network request.

## Telemetry & transparency

`apps/web/src/lib/usage.ts` keeps a ring-buffered ledger (200 entries) in `localStorage`.
`packages/agent/src/orchestrator.ts` attaches an `LlmTelemetry` record - provider, the model the
provider reported serving, prompt/completion/total tokens, latency, and any error - only when the
model layer actually ran. The host turns that into a ledger entry, recording locally answered
turns with zero tokens. This is a diagnostic record, not analytics: nothing is transmitted, and
the user can clear it from the usage page.

## Deployment

`vercel.json` pins the install command, build command, output directory, SPA rewrites,
immutable asset caching, and security headers including a CSP allow-list covering only the
provider and font hosts the app genuinely uses.

## Known follow-ups

- **Routing/SSR** - `@tanstack/react-router` / `@tanstack/react-start` have been removed; the
  app is a client-rendered SPA and needs no router until multi-page navigation is required.
  Reintroduce a router when deep-linkable views (e.g. `/workbooks/:id`) become a real need.
- **Workbook recovery** - the latest successful workbook checkpoint is stored in IndexedDB,
  without chat, keys or undo history. Replacement and edit revisions fence late agent callbacks
  and invalidate old action/plan cards.
- **`xlsx` loading** - the codec is now lazily imported on first upload/export, so the initial
  payload is the app + React chunks only. `loadXlsx()` in
  `apps/web/src/lib/engine-adapter.ts` is the single memoized seam to pre-warm if a future
  feature needs to parse a workbook before the user interacts.
- **Model-layer evals** - provider adapters, native Gemini tool turns, cancellation, full-turn
  usage accounting, and the guardrail's handling of model proposals are covered by recorded-payload
  tests. A live, opt-in contract test against each provider (behind an env flag, never in CI) would
  additionally catch upstream schema drift.
- **Browser E2E** - the jsdom suite renders the real app, but not a real browser. A Playwright
  smoke test (upload, chat, apply, export) would cover paint-level regressions jsdom cannot see.
- **Coverage thresholds** - coverage is not yet enforced in CI. Adding a floor per package would
  make untested additions fail loudly.
- **Cloud memory** - local associations persist in `localStorage`, with verified outcome
  counts restored on startup. Optional Supabase synchronization is enabled only when both
  build-time settings are supplied; queries, operation arguments, and working-step context may
  be sent. Backend provisioning, tenant isolation, access policies, retention/deletion, and the
  backend's exact CSP origin require deployment-specific verification.
