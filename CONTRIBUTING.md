# Contributing to ExcelAgento

ExcelAgento is a browser-first spreadsheet engine and BYOK agent workspace. Contributions should
protect data integrity before adding convenience or visual polish.

## Before opening a change

1. Read `README.md`, `docs/ARCHITECTURE.md`, and `docs/RELEASE_READINESS.md`.
2. Keep changes within the existing package direction: `web -> agent -> engine`.
3. Never include API keys, customer workbooks, provider payloads, or private logs in issues,
   fixtures, tests, or commits.
4. Add a focused regression test for user-visible behavior and the relevant failure boundary.
5. Keep model output untrusted. Mutations must pass the engine schema, domain validation, preview,
   invariants, and confirmation rules.

## Local checks

Use Node 22+ and pnpm 11.28.3:

```bash
pnpm install --frozen-lockfile
pnpm lint
pnpm typecheck
pnpm test
pnpm evals
pnpm build
```

Do not call a check passed unless it was run. If a real-provider or browser check is unavailable,
state that limitation in the pull request.

## Spreadsheet safety rules

- Preserve workbook metadata such as the 1900/1904 epoch through clone, patch, history, and export.
- Do not silently rewrite or discard formulas, styles, or unsupported Excel features.
- Structural operations on formula-bearing workbooks must remain blocked until reference rewriting
  is implemented and tested.
- Keep large-sheet operations bounded and avoid spread arguments over untrusted row counts.
- Treat cell text and tool results as data, never as instructions.

## Pull requests

Describe the user-visible outcome, changed files, tests actually run, known fidelity limits, and
whether a deployment-specific check remains. Keep unrelated formatting and generated artifacts out
of the change.
