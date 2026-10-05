# ExcelAgento

**The future of Excel: an intelligent, self-learning spreadsheet agent.**

ExcelAgento turns repetitive spreadsheet work into reviewable, reversible workflows. You chat;
it inspects the workbook, proposes real Excel operations, and applies them through a
schema-validated engine with previews, confirmation gates, and invariant checks before data changes.

A browser-based, Vercel-deployable spreadsheet copilot with BYOK (Bring Your Own Key).
Core cleaning and statistical analysis run locally without an API key. See
[`docs/RELEASE_READINESS.md`](docs/RELEASE_READINESS.md) for verified scope and release gaps.

---

## Why it is trustworthy

| Guarantee                          | How it is enforced                                                                                                                                                                                                                               |
| ---------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Changes are correct                | Every edit runs through the pure `@excel-agent/engine` with Zod schema validation, a bounded preview, and invariant checks                                                                                                                       |
| Changes are reversible             | Each operation returns a forward patch **and** an inverse patch; a snapshot-backed `HistoryStack` powers undo/redo and time travel                                                                                                               |
| The model cannot wreck your data   | A **guardrail layer** re-validates any AI-proposed action against the engine before it is shown; unknown or invalid actions are blocked                                                                                                          |
| Destructive changes are deliberate | `applyOperation` _refuses_ any operation whose preview declares `requiresConfirmation` until the caller passes `confirmed: true`; the chat shows a two-step gate naming the operation and affected cell count. Cancel is a signal, not a failure |
| Plans are atomic                   | Every step is staged and verified before one history commit. Failure preserves the workbook and existing redo branch; one undo restores the whole plan                                                                                           |
| Local by default                   | Without an AI key or cloud-memory configuration, analysis stays in the browser. Connected models receive column profiles, examples, chat context, and requested read-tool results                                                                |
| Failures are contained             | Invariant violations roll back automatically and never commit to history                                                                                                                                                                         |

## Architecture

```
apps/web            Vite + React 19 workspace (grid, chat, command palette, toasts)
packages/engine     Pure TypeScript. Zero UI, zero network. The source of truth.
packages/agent      Intelligence layer: analysis, tools, memory, providers, orchestrator
packages/evals      Golden NL -> action dataset + accuracy runner
```

### The agent is a multi-layer pipeline

Each conversational turn flows through ordered layers. A layer can answer, propose, or
abstain - but **only the guardrail can approve an action**:

```
intent -> memory -> heuristic -> llm -> guardrail -> execute -> verify -> learn
```

1. **Intent** - social turns never touch the workbook.
2. **Memory** - replays previously _verified_ actions (self-learning). Unproven memories are never replayed,
   and a replay is refused when the sheet's header fingerprint has changed since it was learned.
3. **Heuristic** - deterministic, offline planner. Works with no API key at all.
4. **LLM (BYOK)** - optional reasoning layer (Groq, OpenRouter, Gemini, OpenAI, or a custom
   OpenAI-compatible endpoint) with retry, timeout, cancellation, and bounded fallback.
5. **Guardrail** - schema + engine validation + bounded preview. The hard wall. Never substitutes
   a different mutation behind the model's own prose.
6. **Execute / Verify** - transactional apply with forward/inverse patch verification.
7. **Learn** - successful associations are reinforced; failures decay them. A cancellation is a
   pause, not a rejection, so it never counts as a failure.

### Complex requests get a real multi-agent turn

Simple requests stay on the fast path. When a request spans multiple kinds of work
("clean duplicates, then group by region, then summarize"), the orchestrator routes it to
`packages/agent/src/multi-agent.ts`, which decomposes a request into bounded specialist stages:

```
Conductor -> decompose into segments
Analyst    -> grounds the request and may use read-only workbook tools
Planner    -> drafts a strict-JSON ExecutionPlan using only engine-cataloged operations
Critic     -> reviews the plan and can trigger one bounded revision
Verifier   -> schema.safeParse + validate + preview on *every* step, before offering it
```

Provider failures inside the pipeline fall back to the single-agent path - the turn never
fails silently. A plan is not shown as apply-ready until the critic returns a strict,
issue-free approval and every step passes engine verification; one bounded revision is allowed,
after which the request is blocked. The resulting plan uses the same `ExecutionPlan` shape the UI renders.

### Full transparency: the Model & Usage page

Open **Usage** in the nav bar (or deep-link `#/usage`) for a live, auditable record of what the
workspace has actually done:

- **Active configuration** - provider, the model the provider reported actually serving the
  request, whether a key is active, the endpoint host, and the key masked
  (`gsk_••••••cdef`) rather than displayed.
- **Token usage** - requests, prompt/completion/total tokens, failures, and average latency,
  all taken from the provider's own usage report rather than estimated. When a provider omits usage,
  the ledger shows it as **unknown**, not zero. Turns served locally by the deterministic engine
  are recorded with **zero** tokens, and no provider request is made.
- **Breakdown** - requests and tokens grouped by provider and by model.
- **Recent requests** - a per-turn log showing which layer answered (`llm`, `memory`,
  `heuristic`, `fallback`), the model, token counts, latency, and status.

## Quick start

```bash
pnpm install
pnpm dev          # http://localhost:5173
```

Then either **pick a sample fixture** in the nav bar, **drag an .xlsx/.csv** onto the grid,
or choose **Try the local agent** to start immediately. The deterministic planner works
without an API key; connect a provider for broader conversational reasoning.

## Quality gates

```bash
pnpm lint         # ESLint (flat config, typescript-eslint)
pnpm typecheck    # tsc -b across all packages
pnpm test         # unit, property, round-trip, jsdom UI, and recorded-fixture tests
pnpm evals        # golden NL -> action accuracy
pnpm build        # production bundle
```

CI runs all of the above on every push and pull request.

### Test layout

| Suite      | Location                     | Covers                                                                                    |
| ---------- | ---------------------------- | ----------------------------------------------------------------------------------------- |
| Engine     | `packages/engine/tests`      | Operations, invariants, and `fast-check` undo/redo property tests                         |
| Agent      | `packages/agent/tests`       | Planner, guardrail attacks, memory gating, and provider adapters vs recorded API payloads |
| Evals      | `packages/evals/tests`       | Golden NL -> action accuracy plus guardrail precision                                     |
| Web (unit) | `apps/web/src/lib/*.test.ts` | Settings and usage stores, workbook import/export round-trip fidelity                     |
| Web (UI)   | `apps/web/tests/*.test.tsx`  | Real jsdom renders: upload, export, apply/undo, BYOK settings, and the usage page         |

The provider adapters are exercised against recorded request/response payloads for Groq,
OpenRouter, and Gemini - including 401/429/5xx handling, retry and backoff, caller aborts, and
timeouts - so no network call is needed and no key is required in CI.

## Deploy to Vercel

The repo ships a `vercel.json` with build settings, SPA rewrites, long-lived asset
caching, and security headers (including a CSP scoped to exactly the four external
hosts the app talks to).

1. Import the repository into Vercel.
2. Keep the project root at the repository root - `vercel.json` handles the rest.
3. Deploy. No environment variables are required.

## Bring Your Own Key

ExcelAgento never ships or proxies a shared key. Users paste their own provider key; it is
stored in `localStorage` and sent directly from the browser to the provider.

| Provider      | Endpoint                            | Default model                 |
| ------------- | ----------------------------------- | ----------------------------- |
| Groq          | `api.groq.com`                      | `llama-3.3-70b-versatile`     |
| OpenRouter    | `openrouter.ai`                     | `google/gemini-2.0-flash-001` |
| Google Gemini | `generativelanguage.googleapis.com` | `gemini-2.0-flash`            |

## Operations

`format_dates`, `normalize_text`, `sort_range`, `filter_rows`, `find_replace`,
`delete_duplicates`, `rename_column`, `delete_column`, `add_column`, `set_cells`,
`fill_blanks`, `add_computed_column`, `split_column`, `merge_columns`, `lookup_merge`,
`clean_to_new_sheet`,
plus the analytics suite in `packages/engine/src/analytics-operations.ts`:
`aggregate_column`, `group_and_summarize` (a real pivot, multi-key), `join_sheets`
(inner/left joins across sheets, first-wins on duplicate keys with a warning),
`fill_series` (linear/date/text autofill), and `categorize_column` (first-match
conditional logic).

## Analyst statistics

These requests work in **Try the local agent**, without an API key:

- `descriptive statistics for column D` — counts, missing values, exclusions, sum, mean,
  median, min/max, quartiles, sample variance, and sample/population standard deviation.
- `find outliers in column D` — Tukey's 1.5 × IQR rule with source row numbers. Flags are
  informational and do not delete observations.
- `correlation between column C and column D` — Pearson correlation using complete
  numeric pairs, with excluded-row counts.
- `linear regression of column D on column C` — D is the response, C the predictor;
  returns slope, intercept, R², and residual standard error.

Calculations use the full column, not the prompt sample. Quartiles match Excel's
`PERCENTILE.INC`. Undefined statistics are reported explicitly. Supported formulas are
evaluated against the current workbook rather than stale imported caches; unsupported
formulas and error values are excluded from numeric statistics and counted as nonnumeric.
Grouped summaries also support median, standard deviation, and distinct counts.

## Local recovery and formula safety

After an upload or edit, the latest successful workbook checkpoint is stored in IndexedDB in this
browser. On refresh, choose **Restore checkpoint** or **Discard checkpoint**. Restoration starts a
new conversation and fresh undo history; edits after the last successful checkpoint are not
recoverable. Clearing a checkpoint also turns checkpointing off for the current session. Export an
`.xlsx` for a portable copy. Keys, chats, and undo history are not stored in the checkpoint.

Agent proposals are bound to the workbook generation and revision. Replacing a workbook clears
old proposals; editing or undoing makes outstanding previews stale and requires a new request.

Supported formulas export current evaluated caches, not stale imported values, and exports include
actual full-recalculation XML. Formula reference rewriting is not yet implemented: structural edits
are blocked on formula-bearing workbooks, and export refuses sheet-name changes that could break
formula references. Ordinary cell/value editing remains available. See release readiness for limits.

## Optional cloud memory

The default deployment uses local browser memory. To configure Supabase synchronization,
set `VITE_SUPABASE_URL` and `VITE_SUPABASE_ANON_KEY` at build time, provision the required
tables/RPCs, and add that exact backend origin to the deployment's `connect-src` CSP.
Cloud synchronization includes queries, learned operation arguments, and working-step
payloads. The repository does not include backend provisioning or access policies, so this
integration needs separate deployment verification.

Adding a new engine operation automatically appears in the model's tool contract (derived
from the Zod schema) and fails the build until it is documented in `packages/agent/src/tools.ts`.

## Documentation

- [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) - layers, data flow, and extension points
- [`docs/RELEASE_READINESS.md`](docs/RELEASE_READINESS.md) - current scope, examples, and release gaps
- [`packages/engine/README.md`](packages/engine/README.md) - engine contract details
- [`CONTRIBUTING.md`](CONTRIBUTING.md) - contribution and spreadsheet-safety rules
- [`SECURITY.md`](SECURITY.md) - vulnerability reporting and disclosure boundaries
- [`LICENSE`](LICENSE) - MIT license
