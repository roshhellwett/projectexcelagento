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
