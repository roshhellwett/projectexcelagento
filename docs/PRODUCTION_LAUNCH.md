# ExcelAgento production launch runbook

This runbook is the boundary between a verified repository and a live deployment. It records what
the code checks locally, what the official hosted deployment currently provides, and which actions
still require an operator with access to Vercel, Supabase, and the production account.

## Readiness summary

The repository is ready for a controlled production deployment when the operator completes the
external configuration checklist below. The local product path is covered by deterministic engine
tests, workspace tests, browser tests, privacy-safe release checks, and a production build.

### Verified in the repository

- Plain-English workbook activation flow with sample guidance and an export result state.
- Named, engine-backed workflows for weekly cleanup, date normalization, duplicate removal, and
  sheet reconciliation.
- Preview, confirmation, schema validation, invariant checks, history, and undo boundaries.
- User-visible plan assumptions, safety status, blocked-request explanations, and result receipts.
- Local-first feedback collection that excludes workbook contents, prompts, file names, cell values,
  operation arguments, and provider secrets.
- Open-source contribution, support, security, and issue-routing documentation.
- Production configuration validation and privacy-safe error boundaries.
- Account trial and activation-key entitlement flow backed by Supabase.

### Not claimed by this runbook

- No automated payment checkout is included. The current commercial path is a 30-day evaluation plus
  server-authoritative activation keys managed by the operator console.
- No unattended background job, server-side workbook queue, or model-hosting promise is implied.
- No provider API key, Supabase secret, Vercel credential, domain ownership, or legal approval is
  stored in this repository.

## 1. External configuration gate

Complete these steps in the deployment systems before inviting users:

### Vercel / hosting

Set these **production** environment variables for the web application:

- `VITE_SUPABASE_URL`
- `VITE_SUPABASE_ANON_KEY`
- `VITE_AUTH_REDIRECT_ORIGIN` — the exact HTTPS origin users will visit, without a path or query
  string

Do not put provider model keys, the licensing pepper, Supabase service-role keys, or other private
secrets in `VITE_*` variables. Browser-exposed variables are not secret storage.

Confirm that the production build runs the root command `pnpm build` and publishes `dist`.

### Supabase

1. Apply every migration in `supabase/migrations/` in order.
2. Deploy the `license` Edge Function.
3. Configure the function secrets required by its implementation, including a licensing pepper of at
   least 32 characters and the exact production origin.
4. Set the Auth Site URL to the same production origin.
5. Add the confirmation and recovery callback URLs used by the application.
6. Verify email delivery and password-recovery delivery with a real test account.
7. Enable private vulnerability reporting for the repository or provide the private security channel
   described in `SECURITY.md`.

### Provider access

The deterministic local engine works without a model key. If model-backed reasoning is offered,
each user must either bring a provider key or the operator must explicitly configure a managed
provider path. Verify the provider's terms, quota, region, and data handling before describing it as
available to customers.

## 2. Pre-deploy commands

Run from the repository root on the release commit:

```text
pnpm release:check
pnpm check
pnpm test
pnpm evals
pnpm build
pnpm test:e2e
```

Do not treat a check as passed unless its command completed successfully on the release commit. The
release check must report that no tracked private environment files or deployment credentials were
found.

## 3. Manual smoke test after deployment

Use a fresh test account and a synthetic workbook; do not use a customer workbook for smoke tests.

### Account and access

- [ ] Sign up with a test email.
- [ ] Confirm the email through the newest confirmation link.
- [ ] Sign in and confirm the account shows an active evaluation.
- [ ] Refresh the page and confirm access remains active.
- [ ] Request password recovery and complete it with the newest link.
- [ ] Activate a generated test key and confirm the entitlement and expiry date update.
- [ ] Revoke or expire the test entitlement in the operator console and confirm the workspace locks.

### First verified workbook result

- [ ] Open the workspace with no model key.
- [ ] Use the sample workbook or upload a synthetic `.xlsx` file.
- [ ] Choose a starter workflow or enter a plain-English request.
- [ ] Confirm the preview identifies the affected range and assumptions.
- [ ] Verify that no workbook mutation occurs before Apply.
- [ ] Apply the request and verify the result receipt.
- [ ] Export the workbook and open the downloaded file in a spreadsheet application.
- [ ] Undo the committed change and verify the original workbook state returns.

### Safety, privacy, and failure states

- [ ] Submit an unsupported request and confirm it is blocked without an Apply action.
- [ ] Change the workbook after planning and confirm the stale plan must be replanned.
- [ ] Trigger the error boundary with a test-only failure and confirm the UI shows only a reference
      code, not a raw exception or component stack.
- [ ] Open Model & Usage and confirm local usage/feedback controls work.
- [ ] If using BYOK, confirm the Privacy page matches the provider and payload behavior being offered.

## 4. Commercial access model

The current production-ready monetization boundary is intentionally simple:

1. The open-source core provides the local-first workbook engine and reviewable workflow.
2. The official hosted deployment requires an account and starts a 30-day evaluation.
3. An operator issues activation keys for paid or extended access through the admin console.
4. The server validates and records entitlements; the browser never decides whether an expired key
   is valid.
5. Provider model costs remain separate when users bring their own key.

This supports a controlled launch without inventing a price or integrating a payment processor before
the business chooses one. A future checkout integration must preserve the same server-authoritative
entitlement boundary and should be added only with a selected provider, refund policy, tax treatment,
and webhook/reconciliation plan.

## 5. Launch day operations

### Before announcing the URL

- [ ] Record the deployment URL and release commit in the operator's private release log.
- [ ] Confirm the production Supabase project and Vercel project are the intended projects.
- [ ] Confirm no staging or local callback URL remains in Auth configuration.
- [ ] Create one operator account and one non-operator test account; never use the operator account
      for ordinary user testing.
- [ ] Generate only the activation keys needed for the first cohort and store plaintext keys in the
      operator's private channel. The application stores server-side keyed hashes for verification;
      generated plaintext keys are shown once and are not persisted by the hardened migration.
- [ ] Publish the support and security routes that are actually monitored.
- [ ] Keep the first design-partner cohort small enough that workbook incidents can be handled
      manually.

### If a deployment fails

1. Check the browser-safe configuration error and the deployment logs; do not ask users to paste API
   keys into chat or issue trackers.
2. If auth is unavailable, verify the three `VITE_*` values and Supabase callback configuration.
3. If access is locked, inspect the account and license event in the operator console before changing
   an entitlement.
4. If a workbook task fails, preserve the mission receipt and the reference code, then reproduce
   with a synthetic fixture.
5. If a model provider fails, route users to deterministic workflows or their configured provider;
   do not silently claim that a model-backed result was produced.
6. Roll back the hosting deployment only after confirming whether the issue is in the frontend build,
   Supabase configuration, or an external provider.

## 6. Human sign-off

The repository can prove code behavior, but only the deployment operator can sign off these facts:

- [ ] Production secrets and callback URLs are configured correctly.
- [ ] Supabase migrations and the license function are deployed.
- [ ] Email confirmation and recovery work from the public origin.
- [ ] Support and security requests reach a monitored channel.
- [ ] The current evaluation/activation terms, refunds, taxes, and customer communications are
      approved by the business owner.
- [ ] Privacy, terms, and any jurisdiction-specific legal review are approved for the actual
      operator and deployment.
- [ ] A named person owns incident response for the first cohort.

Only after these boxes are checked should the deployment be described as production-ready to users.
