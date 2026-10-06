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

## Email confirmation and password recovery

Signup and verification resend explicitly use the app's current origin and base path with
`?auth=confirm`. Password recovery uses the same destination with `?auth=recovery`.
For example, an app hosted at `https://your-app.example/` sends users back to:

- `https://your-app.example/?auth=confirm`
- `https://your-app.example/?auth=recovery`

The Supabase SDK consumes the email callback before the app clears its token fragment. The
return screen offers workspace access only after a session is established. Recovery links
open a new-password form; expired links offer sign-in and a fresh email request.

### Required Supabase dashboard settings

In **Authentication → URL Configuration** for the Supabase project used by this deployment:

1. Set **Site URL** to the actual production app URL, including `https://`. Do not leave the
   production fallback set to `http://localhost:3000`.
2. Add both callback URLs above to **Redirect URLs**. Include any base path if the app is
   hosted below `/app/` or another prefix.
3. For local development, also allow `http://localhost:5173/?auth=confirm` and
   `http://localhost:5173/?auth=recovery`. Add the `127.0.0.1` equivalents if used.
4. In **Authentication → Email Templates**, confirmation and password-reset links should use
   `{{ .ConfirmationURL }}`. A link hardcoded to localhost or built only from `{{ .SiteURL }}`
   discards the per-request destination. The app uses Supabase's default session-producing
   email links, not a custom `token_hash` verification endpoint.

Supabase can fall back to **Site URL** when a requested destination is not allowlisted. These
dashboard settings must be updated separately from the application deployment.

### Recovering an old link that opened localhost

First return to the deployed app and try **Sign in**: the verification may have succeeded
before the redirect failed. If the account is still unconfirmed, use **Resend confirmation
email** and enter the existing account email. Do not create a second account. Existing emails
keep their original destination; use the newest link after the settings above are corrected.
The app waits 60 seconds between successful email requests and surfaces provider rate limits.
