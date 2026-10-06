# ExcelAgento licensing operations

ExcelAgento's official deployment uses a server-authoritative activation service. The browser
never receives a service-role key, raw activation keys are not stored in Postgres, and browser
storage is treated as an installation signal rather than an immutable hardware identity.

## Components

- `supabase/migrations/20261006000000_create_license_system.sql` creates the trial, device,
  license, admin allowlist, and audit tables.
- `supabase/functions/license/index.ts` is the authenticated API for account status, trial
  creation, activation, heartbeat, key generation, day adjustments, revocation, banning, and
  ownership-verified transfers.
- `apps/web/src/lib/device-identity.ts` creates a stable per-browser-install UUID. It does not
  collect WebGL, canvas, font, or other invasive fingerprints.
- `apps/web/src/components/LicensePanel.tsx` shows the signed-in email, installation ID, trial
  or license state, expiry, activation form, and recovery contact.
- `apps/web/src/components/AdminLicensePage.tsx` is the authenticated operations console.

## Supabase deployment

Apply the migration with the Supabase CLI or the SQL editor, then deploy the Edge Function:

```bash
supabase db push
supabase functions deploy license --project-ref fsepapdadtrlddkyqqxu
```

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

`APP_ORIGIN` should be the exact production origin, including no trailing slash. For local
development, use a separate Edge Function environment with `APP_ORIGIN=http://localhost:5173`.

## Bootstrap the first administrator

The migration intentionally does not guess an administrator email. After creating and confirming
the operator's Supabase account, insert that user's UUID in the SQL editor:

```sql
insert into public.license_admins (user_id, role)
select id, 'owner'
from auth.users
where lower(email) = lower('zenithopensourceprojects@gmail.com')
on conflict (user_id) do update set role = excluded.role;
```

Only users in `license_admins` can call administrative actions. The frontend does not contain an
admin password or an authorization secret.

## Operational rules

- A confirmed account receives one 30-day trial when the licensing service first sees it.
- A license key is generated once, stored as a salted hash, and returned in plaintext only in the
  administrator's generation response. If that response is lost, generate a replacement key.
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

The system protects the official hosted deployment. Because the application is open source and
browser-delivered, a third-party fork can remove client-side gating. Server checks therefore remain
the source of truth for the official service, while the UI communicates the limitation clearly.
