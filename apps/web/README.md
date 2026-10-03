# @excel-agent/web

The ExcelAgento workspace UI: a Vite + React 19 single-page application.

| File                        | Responsibility                                                                                 |
| --------------------------- | ---------------------------------------------------------------------------------------------- |
| `src/App.tsx`               | Workspace shell, transactional operation execution, self-learning feedback loop                |
| `src/components/*`          | Grid, chat, command palette, history drawer, settings, toasts, error boundary                  |
| `src/lib/engine-adapter.ts` | `.xlsx`/`.csv` import and export (Excel-legal sheet names, number formats, blank preservation) |
| `src/lib/agent-runtime.ts`  | Shared registry, self-learning memory, and orchestrator instances                              |
| `src/lib/settings.ts`       | BYOK settings store with legacy-key migration                                                  |
| `src/lib/llm-service.ts`    | Thin adapter over the agent orchestrator                                                       |

Run `pnpm dev` from the repository root. Builds to `dist/` and deploys through the root
`vercel.json`.
