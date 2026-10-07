# ExcelAgento licensing operations

ExcelAgento's official deployment uses a server-authoritative activation service. The browser
never receives a service-role key. Generated admin keys are shown in plaintext only in the
generation response and in the current admin page session; older hash-only keys cannot be
recovered and must be replaced. The browser installation signal is a best-effort device
fingerprint, not cryptographic proof of a physical machine.

## Components

- `supabase/migrations/20261006000000_create_license_system.sql` creates the trial, device,
  license, admin allowlist, and audit tables.
- `supabase/functions/license/index.ts` is the authenticated API for account status, trial
  creation, activation, heartbeat, key generation, day adjustments, revocation, banning, and
  ownership-verified transfers.
- `apps/web/src/lib/device-identity.ts` derives a UUID from WebGL renderer/vendor, CPU cores, and
  display characteristics, then persists it locally. Fingerprints can collide or change when
  graphics privacy settings, hardware, or browser storage changes; support recovery remains
  necessary.
- `apps/web/src/components/LicensePanel.tsx` shows the signed-in email, installation ID, trial
  or license state, expiry, activation form, and recovery contact.
- `apps/web/src/components/AdminLicensePage.tsx` is the authenticated operations console.

## Supabase deployment

Apply all migrations with the Supabase CLI or the SQL editor, then deploy the Edge Function:

```bash
supabase db push
supabase functions deploy license --project-ref fsepapdadtrlddkyqqxu
```

If `20261006000000_create_license_system.sql` was already run manually, run the upgrade migrations
in timestamp order before redeploying the function. The latest guard prevents a device collision or
shared-network threshold from permanently banning the device/account automatically, and it
bootstraps the configured owner email as an `owner` admin after email confirmation.
It fixes active-key device reset/rebinding, adds shared-read verification concurrency, and preserves
the existing administrative records.

Supabase injects `SUPABASE_URL` and `SUPABASE_SERVICE_ROLE_KEY` into deployed Edge Functions.
Set the remaining deployment secrets with the Supabase secret store. Do not put any of them in
`VITE_*` variables or the browser:

```text
APP_ORIGIN=https://excelagento.vercel.app
LICENSE_KEY_PEPPER=<long-random-secret>
LICENSE_DEVICE_PEPPER=<different-long-random-secret>
```

From a shell with OpenSSL available:

```bash
supabase secrets set \
  APP_ORIGIN=https://excelagento.vercel.app \
  LICENSE_KEY_PEPPER="$(openssl rand -hex 32)" \
  LICENSE_DEVICE_PEPPER="$(openssl rand -hex 32)" \
  --project-ref fsepapdadtrlddkyqqxu
```

`APP_ORIGIN` should be the exact production origin, including no trailing slash. The web deployment
also needs `VITE_AUTH_REDIRECT_ORIGIN=https://excelagento.vercel.app` so Supabase confirmation and
recovery requests never inherit a localhost or preview origin. For local
development, use a separate Edge Function environment with `APP_ORIGIN=http://localhost:5173`.

## Bootstrap the first administrator

The migration intentionally does not guess an administrator email. After creating and confirming
the operator's Supabase account, insert that user's UUID in the SQL editor:

```sql
insert into public.license_admins (user_id, role)
select id, 'owner'
from auth.users
where lower(email) = lower('zenithprojects@icloud.com')
on conflict (user_id) do update set role = excluded.role;
```

Only users in `license_admins` can call administrative actions. The frontend does not contain an
admin password or an authorization secret.

## Operational rules

- A confirmed account receives one 30-day trial when the licensing service first sees it.
- A license key is generated once and returned in plaintext only in the administrator's generation
  response/current admin session. If that response is lost, generate a replacement key; the UI
  never copies a masked hint as though it were a usable key.
- Activation binds the key to the Supabase user ID and the current browser installation record.
- A device transfer can either preserve the verified device binding or clear it for activation on
  a new installation. The admin must supply an ownership-verification note of at least 12
  characters; that note is retained in the audit event.
- Revocation, expiry, device reset, ban, unban, adjustment, and transfer are audited.
- The workspace fails closed when the server cannot verify an authenticated account. This avoids
  treating editable browser storage as proof of entitlement.
- The installation ID can change when browser storage is cleared, a different browser is used, or
  private browsing is used. Support should use the signed-in email, displayed installation ID, and
  audit history to recover legitimate users.

## Fresh-launch reset procedure

The repository does not run a destructive Supabase reset automatically. Before launch, perform the
reset against the confirmed project ref `fsepapdadtrlddkyqqxu` in the Supabase dashboard or an
authenticated Supabase CLI session:

1. Export a database backup and record the owner user's Auth UUID.
2. Confirm `zenithprojects@icloud.com` exists and is email-confirmed. Do not delete this
   user.
3. Remove non-owner Auth users from **Authentication → Users** and clear application tables owned by
   the project. Review any tables outside the licensing migrations before truncating them.
4. Clear licensing rows in dependency order: events/rate limits, keys, devices, accounts, and
   admin rows. Keep or recreate the owner admin row.
5. Apply every migration in timestamp order, including
   `20261007120000_guard_license_false_positive_protection.sql`.
6. Deploy the `license` Edge Function and set `APP_ORIGIN`, both peppers, and the production auth
   redirect environment variable.
7. Test owner status, a new confirmed-user trial, a second account on the same device, an expired
   trial, key generation/copy, and an email-confirmation link before opening the site publicly.

Never run a blanket `drop schema public cascade` or delete `auth.users` without first verifying the
owner UUID and the target Supabase project. The application cannot safely infer those details from
the browser.

The system protects the official hosted deployment. Because the application is open source and
browser-delivered, a third-party fork can remove client-side gating. Server checks therefore remain
the source of truth for the official service, while the UI communicates the limitation clearly.
