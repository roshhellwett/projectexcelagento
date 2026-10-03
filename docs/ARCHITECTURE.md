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
AgentDecision { message, action?, guardrail?, source, trace }
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

## Extension points

| Goal                           | Where to change                                                                                                                   |
| ------------------------------ | --------------------------------------------------------------------------------------------------------------------------------- |
| Add an Excel operation         | `packages/engine/src/operations.ts`, then register it and document it in `packages/agent/src/tools.ts` (build fails until you do) |
| Support a new model provider   | `packages/agent/src/providers.ts` - implement `ProviderAdapter`, add to `ADAPTERS`                                                |
| Teach the agent a new phrasing | `packages/agent/src/analysis.ts`, then add an eval case in `packages/evals/src/cases.ts`                                          |
| Change the UI shell            | `apps/web/src/components/*` - components are presentational and receive props                                                     |

## Testing strategy

- **Engine** - unit tests per operation plus `fast-check` property tests for undo/redo.
- **Agent** - guardrail semantics, memory gating, tool-catalog/registry parity.
- **Evals** - a golden set of natural-language prompts mapped to expected operations,
  asserting both accuracy and that 100% of proposed actions clear the guardrail.
- **Web** - workbook import/export round-trip fidelity (values, formulas, blanks, number
  formats, sheet-name sanitization, CSV detection).

## Deployment

`vercel.json` pins the install command, build command, output directory, SPA rewrites,
immutable asset caching, and security headers including a CSP allow-list covering only the
provider and font hosts the app genuinely uses.

## Known follow-ups

- `@tanstack/react-router` / `@tanstack/react-start` are declared but unused while the app
  is a client-rendered SPA. They are tree-shaken out of the bundle; either adopt them for
  routing/SSR or remove them.
- `xlsx` is isolated in its own chunk. Lazy-loading it on first upload would shrink the
  initial payload further.
