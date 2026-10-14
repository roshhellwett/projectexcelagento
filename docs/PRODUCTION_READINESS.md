# Production readiness

The web build fails closed when a production deployment does not have explicit public
authentication configuration. The startup screen lists missing variable names without echoing their
values, and the SDK is pointed at a disabled placeholder until the deployment is rebuilt.

This page documents the application boundary. [`PRODUCTION_LAUNCH.md`](./PRODUCTION_LAUNCH.md) is
the operator runbook for Supabase migrations, the licensing Edge Function, post-deploy smoke tests,
commercial access sign-off, and incident handling.

## Build-time configuration

Set these variables in the Vercel project (or the equivalent build environment) before a production
build:

- `VITE_SUPABASE_URL` — the HTTPS Supabase project URL.
- `VITE_SUPABASE_ANON_KEY` — the public Supabase client key. Keep it in the deployment environment;
  do not add private provider keys to Vercel configuration files or the repository.
- `VITE_AUTH_REDIRECT_ORIGIN` — the HTTPS origin that Supabase email confirmation and recovery
  links should return to.

Local development and tests retain their existing browser-local fallbacks. Production builds require
all three values and reject malformed or non-HTTPS auth URLs before mounting the app.

## Checks

Run these from the repository root before a release:

```text
pnpm release:check
pnpm validate
pnpm build
```

`release:check` is side-effect-free. It verifies the tracked-file boundary, checks both Vercel
configuration files for credential-shaped values, confirms that the example environment file
documents the required variable names, and rejects direct non-empty fallback literals for the
licensing Edge Function's origin and key pepper. It is intentionally a limited static hygiene check,
not a complete secret scanner or proof that a hosted service is configured. CI runs it before the
normal formatting, lint, typecheck, test, build, and browser checks.

The licensing function also fails closed unless `APP_ORIGIN`, `LICENSE_KEY_PEPPER` (at least 32
characters), `SUPABASE_URL`, and `SUPABASE_SERVICE_ROLE_KEY` are present in its private Edge
Function environment. Unexpected licensing failures return a bounded reference code rather than
echoing provider, database, or exception details to the browser.

## Human / external actions still required

- Set the three build variables in the deployment provider without committing their values.
- Set the licensing function's private environment values and deploy every migration in
  `supabase/migrations/` before offering account activation.
- Configure the same production origin and callback URLs in the Supabase Authentication URL
  Configuration for the project used by the deployment.
- Confirm provider API keys, model access, quotas, and any custom endpoint availability in the
  user's own provider accounts; this repository does not test or provision those services.
- Review the deployed startup screen and complete a synthetic sign-up, confirmation, sign-in, and
  password-recovery smoke test after deployment.
