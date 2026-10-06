export interface LicenseDependencies {
  origin: string;
  keyPepper: string;
  devicePepper: string;
  authenticate: (token: string) => Promise<string | null>;
  command: (
    actorId: string,
    action: string,
    payload: Record<string, unknown>,
  ) => Promise<Record<string, unknown>>;
}

const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const KEY = /^EXCEL(?:-[0-9A-F]{8}){4}$/;
const ACTIONS = new Set([
  'status',
  'heartbeat',
  'activate',
  'admin_list',
  'admin_generate',
  'admin_adjust_days',
  'admin_set_duration',
  'admin_set_expiry',
  'admin_revoke',
  'admin_reset_device',
  'admin_transfer',
  'admin_ban_user',
  'admin_unban_user',
  'admin_ban_device',
  'admin_unban_device',
  'admin_adjust_trial',
]);

async function hash(value: string, pepper: string, domain: string): Promise<string> {
  const key = await crypto.subtle.importKey(
    'raw',
    new TextEncoder().encode(pepper),
    { name: 'HMAC', hash: 'SHA-256' },
    false,
    ['sign'],
  );
  const result = await crypto.subtle.sign(
    'HMAC',
    key,
    new TextEncoder().encode(`${domain}:${value}`),
  );
  return Array.from(new Uint8Array(result), (byte) => byte.toString(16).padStart(2, '0')).join('');
}

function newKey(): string {
  const hex = Array.from(crypto.getRandomValues(new Uint8Array(16)), (byte) =>
    byte.toString(16).padStart(2, '0'),
  )
    .join('')
    .toUpperCase();
  return `EXCEL-${hex.match(/.{8}/g)!.join('-')}`;
}

export function createLicenseHandler(
  deps: LicenseDependencies,
): (request: Request) => Promise<Response> {
  const origin = deps.origin.replace(/\/$/, '');
  const respond = (body: Record<string, unknown>, status = 200) =>
    new Response(JSON.stringify(body), {
      status,
      headers: {
        'Content-Type': 'application/json',
        'Cache-Control': 'no-store',
        Vary: 'Origin',
        'Access-Control-Allow-Origin': origin,
        'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
        'Access-Control-Allow-Methods': 'POST, OPTIONS',
      },
    });
  return async (request) => {
    if (!origin || deps.keyPepper.length < 32 || deps.devicePepper.length < 32)
      return respond(
        { error: 'Licensing service is not configured.', code: 'SERVICE_NOT_CONFIGURED' },
        503,
      );
    if (request.headers.get('Origin') && request.headers.get('Origin') !== origin)
      return respond({ error: 'Origin is not allowed.', code: 'ORIGIN_NOT_ALLOWED' }, 403);
    if (request.method === 'OPTIONS')
      return new Response(null, { status: 204, headers: respond({}).headers });
    if (request.method !== 'POST') return respond({ error: 'Method not allowed.' }, 405);
    try {
      const bearer = /^Bearer\s+(\S+)$/i.exec(request.headers.get('Authorization') ?? '');
      const actorId = bearer ? await deps.authenticate(bearer[1]!) : null;
      if (!actorId)
        return respond({ error: 'Sign in again to verify access.', code: 'AUTH_REQUIRED' }, 401);
      const text = await request.text();
      if (text.length > 8192) return respond({ error: 'Request is too large.' }, 413);
      let body: Record<string, unknown>;
      try {
        const parsed: unknown = JSON.parse(text);
        if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed))
          throw new Error('Invalid body');
        body = parsed as Record<string, unknown>;
      } catch {
        return respond({ error: 'Request must be a JSON object.' }, 400);
      }
      const action = typeof body.action === 'string' ? body.action : '';
      if (!ACTIONS.has(action)) return respond({ error: 'Unknown licensing action.' }, 400);
      const payload: Record<string, unknown> = {};
      for (const field of [
        'licenseId',
        'userId',
        'deviceId',
        'targetEmail',
        'verificationNote',
        'reason',
        'deltaDays',
        'durationDays',
        'expiresAt',
        'preserveDevice',
        'offset',
        'query',
      ]) {
        if (body[field] !== undefined) payload[field] = body[field];
      }
      const installId =
        typeof body.installId === 'string' ? body.installId.trim().toLowerCase() : '';
      if (!action.startsWith('admin_') || installId) {
        if (!UUID.test(installId))
          return respond(
            { error: 'A valid installation ID is required.', code: 'INSTALL_ID_REQUIRED' },
            400,
          );
        payload.installHash = await hash(installId, deps.devicePepper, 'excelagento:install:v1');
        payload.installHint = `${installId.slice(0, 8)}…${installId.slice(-6)}`;
      }
      if (action === 'activate') {
        const licenseKey =
          typeof body.licenseKey === 'string' ? body.licenseKey.trim().toUpperCase() : '';
        if (!KEY.test(licenseKey))
          return respond(
            { error: 'Enter a valid EXCEL activation key.', code: 'INVALID_KEY' },
            400,
          );
        payload.keyHash = await hash(licenseKey, deps.keyPepper, 'excelagento:key:v1');
      }
      const keys: string[] = [];
      if (action === 'admin_generate') {
        const count = body.count;
        if (typeof count !== 'number' || !Number.isInteger(count) || count < 1 || count > 100)
          return respond({ error: 'Key count must be an integer between 1 and 100.' }, 400);
        for (let index = 0; index < count; index += 1) keys.push(newKey());
        payload.keys = await Promise.all(
          keys.map(async (key) => ({
            hash: await hash(key, deps.keyPepper, 'excelagento:key:v1'),
            hint: `EXCEL-••••-${key.slice(-6)}`,
          })),
        );
      }
      const result = await deps.command(actorId, action, payload);
      if (!result || typeof result !== 'object') throw new Error('Invalid database response');
      if (result.error) return respond(result, Number(result.httpStatus) || 400);
      return respond({
        ...result,
        ...(keys.length ? { keys } : {}),
        ...(!action.startsWith('admin_') ? { deviceId: installId } : {}),
      });
    } catch {
      return respond(
        {
          error: 'The activation service could not verify access. Please retry.',
          code: 'SERVICE_UNAVAILABLE',
        },
        503,
      );
    }
  };
}
