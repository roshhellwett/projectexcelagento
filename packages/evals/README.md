# @excel-agent/evals

A golden dataset that measures ExcelAgento's natural-language planner and locks in
regression behaviour. It runs entirely offline - no API key required.

- `src/cases.ts` - labelled prompts mapped to expected operations (and optional arg subsets)
- `src/fixtures.ts` - the canonical worksheet every case runs against
- `src/runner.ts` - resolver adapter, scoring, and a printable summary

```ts
import { createOperationRegistry } from '@excel-agent/engine';
import { createOrchestrator } from '@excel-agent/agent';
import { createOrchestratorResolver, formatSummary, runEvals } from '@excel-agent/evals';

const orchestrator = createOrchestrator({ registry: createOperationRegistry() });
const summary = await runEvals(createOrchestratorResolver(orchestrator));

console.log(formatSummary(summary));
```

Two numbers are reported:

- **accuracy** - how often the planner picks the expected operation (and matching args).
- **action precision** - the share of proposed actions that clear the guardrail. This must
  stay at `1.0`: the agent should never offer an action the engine would reject.

Run with `pnpm evals` from the repository root. CI enforces the baseline.

## Coverage

Cases are grouped by intent family and deliberately include Hinglish/slang phrasings and
misspellings, because those are what real users type:

| Family             | Covered                                                                  |
| ------------------ | ------------------------------------------------------------------------ |
| Deduplication      | explicit, slang, Hinglish, and the misspelling "dublicate"               |
| Text normalization | trim, upper/lower/title case, addressed by letter or header name         |
| Dates              | ISO/US/EU targets, Hinglish, and a locale word-boundary regression       |
| Sorting            | asc/desc, "highest first", "order by", "z-a", Hinglish                   |
| Columns            | delete/rename by letter or header name, header-casing regression         |
| Find & replace     | "replace X with Y" and "find X and replace with Y"                       |
| Filtering          | numeric gt/lt (including Hinglish "kam") and categorical equals          |
| Read-only turns    | aggregations, distributions, missing-value audits, overviews, small talk |

Several cases exist specifically as regression locks for planner bugs that the suite caught:
bare-letter column resolution (`"A"` must not match the `a` in `Amount`), locale detection
(`"must"` contains `us`), replacement capture greediness, header casing preservation, and
structural commands outranking keyword heuristics.

`add_column` and `set_cells` are reachable through the LLM tool layer but are not proposed by
the deterministic planner, so they are intentionally absent from this suite.
