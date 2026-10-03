# ExcelAgento

**The future of Excel: an intelligent, self-learning spreadsheet agent.**

ExcelAgento turns hours of spreadsheet work into minutes. You chat; it plans, validates,
and applies real Excel operations - deterministically, reversibly, and with every change
verified against invariants before it touches your data.

Production-ready, Vercel-deployable, and BYOK (Bring Your Own Key): no server, no
uploaded spreadsheets, no secrets to manage.

---

## Why it is trustworthy

| Guarantee                        | How it is enforced                                                                                                                      |
| -------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------- |
| Changes are correct              | Every edit runs through the pure `@excel-agent/engine` with Zod schema validation, a bounded preview, and invariant checks              |
| Changes are reversible           | Each operation returns a forward patch **and** an inverse patch; a snapshot-backed `HistoryStack` powers undo/redo and time travel      |
| The model cannot wreck your data | A **guardrail layer** re-validates any AI-proposed action against the engine before it is shown; unknown or invalid actions are blocked |
| Your data stays local            | Spreadsheets never leave the browser. Only a compact column profile is sent to your chosen model                                        |
| Failures are contained           | Invariant violations roll back automatically and never commit to history                                                                |

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
2. **Memory** - replays previously _verified_ actions (self-learning). Unproven memories are never replayed.
3. **Heuristic** - deterministic, offline planner. Works with no API key at all.
4. **LLM (BYOK)** - optional reasoning layer (Groq, OpenRouter, Gemini) with retry, timeout, and backoff.
5. **Guardrail** - schema + engine validation + bounded preview. The hard wall.
6. **Execute / Verify** - transactional apply with forward/inverse patch verification.
7. **Learn** - successful associations are reinforced; failures decay them.

### Full transparency: the Model & Usage page

Open **Usage** in the nav bar (or deep-link `#/usage`) for a live, auditable record of what the
workspace has actually done:

- **Active configuration** - provider, the model the provider reported actually serving the
  request, whether a key is active, the endpoint host, and the key masked
  (`gsk_••••••cdef`) rather than displayed.
- **Token usage** - requests, prompt/completion/total tokens, failures, and average latency,
  all taken from the provider's own usage report rather than estimated. Turns served locally by
  the deterministic engine are recorded with **zero** tokens, so the ledger doubles as proof that
  nothing left the browser.
- **Breakdown** - requests and tokens grouped by provider and by model.
- **Recent requests** - a per-turn log showing which layer answered (`llm`, `memory`,
  `heuristic`, `fallback`), the model, token counts, latency, and status.

## Quick start

```bash
pnpm install
pnpm dev          # http://localhost:5173
```

Then either **pick a sample fixture** in the nav bar, **drag an .xlsx/.csv** onto the grid,
or paste a BYOK key to unlock the conversational agent. The deterministic planner works
with no key at all.

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
`delete_duplicates`, `rename_column`, `delete_column`, `add_column`, `set_cells`.

Adding a new engine operation automatically appears in the model's tool contract (derived
from the Zod schema) and fails the build until it is documented in `packages/agent/src/tools.ts`.

## Documentation

- [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) - layers, data flow, and extension points
- [`packages/engine/README.md`](packages/engine/README.md) - engine contract details
