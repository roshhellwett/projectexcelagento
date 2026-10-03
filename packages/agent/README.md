# @excel-agent/agent

The intelligence layer. UI-free, network-capable, deterministic by default.

| Module            | Responsibility                                                                               |
| ----------------- | -------------------------------------------------------------------------------------------- |
| `analysis.ts`     | Column profiling, fuzzy cell search, column resolution, and the deterministic intent planner |
| `context.ts`      | Token-efficient sheet context, system prompt, and tolerant model-output parsing              |
| `tools.ts`        | Tool catalog derived **live** from the engine registry (Zod -> JSON Schema)                  |
| `memory.ts`       | Self-learning store of verified `(query -> operation)` associations                          |
| `providers.ts`    | BYOK adapters for Groq, OpenRouter, and Gemini with retry, timeout, and backoff              |
| `orchestrator.ts` | The ordered multi-layer control plane with a hard guardrail                                  |

```ts
import { createOperationRegistry } from '@excel-agent/engine';
import { createMemoryStore, createOrchestrator } from '@excel-agent/agent';

const orchestrator = createOrchestrator({
  registry: createOperationRegistry(),
  memory: createMemoryStore(),
});

const decision = await orchestrator.decide({
  query: 'remove duplicate rows',
  workbook,
  sheetName: 'Orders',
});

if (decision.action && decision.guardrail?.passed) {
  // hand decision.action to engine.applyOperation()
  orchestrator.learn({
    query,
    sheetName: 'Orders',
    operation: decision.action.name,
    args: decision.action.args,
    success: true,
  });
}
```

`buildToolCatalog` throws `ToolCatalogError` if the engine exposes an operation that is not
documented here, so the model contract cannot silently drift from the engine.
