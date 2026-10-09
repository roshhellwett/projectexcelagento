import { describe, expect, it, vi } from 'vitest';
import { createLicenseHandler, type LicenseDependencies } from './handler.js';

const APP_ORIGIN = 'https://excelagento.example';
const PEPPER = '0123456789abcdef0123456789abcdef';

function setup(overrides: Partial<LicenseDependencies> = {}) {
  const dependencies: LicenseDependencies = {
    origin: APP_ORIGIN,
    keyPepper: PEPPER,
    authenticate: vi.fn().mockResolvedValue('actor-id'),
    command: vi.fn().mockResolvedValue({ state: 'trial', canUse: true }),
    ...overrides,
  };
  return { dependencies, handler: createLicenseHandler(dependencies) };
}

async function json(response: Response): Promise<Record<string, unknown>> {
  return (await response.json()) as Record<string, unknown>;
}

function request(body: Record<string, unknown>, headers: Record<string, string> = {}) {
  return new Request(`${APP_ORIGIN}/functions/v1/license`, {
    method: 'POST',
    headers: {
      Origin: APP_ORIGIN,
      Authorization: 'Bearer test-token',
      'Content-Type': 'application/json',
      ...headers,
    },
    body: JSON.stringify(body),
  });
}

describe('license Edge Function boundary', () => {
  it('fails closed when deployment secrets are incomplete', async () => {
    const { handler } = setup({ keyPepper: 'short' });
    const response = await handler(request({ action: 'status' }));
    expect(response.status).toBe(503);
    expect((await json(response)).code).toBe('SERVICE_NOT_CONFIGURED');
  });

  it('rejects missing or invalid authentication before calling the database command', async () => {
    const { dependencies, handler } = setup({ authenticate: vi.fn().mockResolvedValue(null) });
    const response = await handler(request({ action: 'status' }));
    expect(response.status).toBe(401);
    expect(dependencies.command).not.toHaveBeenCalled();
  });

  it('rejects a foreign origin and invalid action', async () => {
    const originSetup = setup();
    const foreign = await originSetup.handler(
      request({ action: 'status' }, { Origin: 'https://attacker.example' }),
    );
    expect(foreign.status).toBe(403);

    const invalidAction = await originSetup.handler(request({ action: 'hack_the_planet' }));
    expect(invalidAction.status).toBe(400);
    expect((await json(invalidAction)).code).toBe('INVALID_ACTION');
  });

  it('passes activation key material to the transactional command without device ID', async () => {
    const { dependencies, handler } = setup();
    const rawKey = 'EXCEL-01234567-89ABCDEF-01234567-89ABCDEF';
    const response = await handler(request({ action: 'activate', licenseKey: rawKey }));
    expect(response.status).toBe(200);
    const [, action, payload] = (dependencies.command as ReturnType<typeof vi.fn>).mock
      .calls[0] as [string, string, Record<string, unknown>];
    expect(action).toBe('activate');
    expect(payload.rawKey).toBe(rawKey);
    expect(payload.keyHash).toMatch(/^[0-9a-f]{64}$/);
    expect(payload).not.toHaveProperty('installId');
    expect(payload).not.toHaveProperty('installHash');
  });

  it('rejects invalid activation key format', async () => {
    const { dependencies, handler } = setup();
    const response = await handler(request({ action: 'activate', licenseKey: 'invalid-key' }));
    expect(response.status).toBe(400);
    expect((await json(response)).code).toBe('INVALID_KEY');
    expect(dependencies.command).not.toHaveBeenCalled();
  });

  it('generates bounded batches and returns plaintext keys only in that response', async () => {
    const { dependencies, handler } = setup();
    const response = await handler(
      request({ action: 'admin_generate', count: 3, durationDays: 60 }),
    );
    const body = await json(response);
    const generated = body.keys as string[];
    expect(response.status).toBe(200);
    expect(generated).toHaveLength(3);
    const [, action, payload] = (dependencies.command as ReturnType<typeof vi.fn>).mock
      .calls[0] as [string, string, Record<string, unknown>];
    expect(action).toBe('admin_generate');
    expect(payload.keys).toHaveLength(3);
    expect((payload.keys as Array<Record<string, unknown>>)[0]?.rawKey).toBe(generated[0]);
  });

  it('rejects unbounded administrator key requests', async () => {
    const { dependencies, handler } = setup();
    const response = await handler(
      request({ action: 'admin_generate', count: 101, durationDays: 30 }),
    );
    expect(response.status).toBe(400);
    expect(dependencies.command).not.toHaveBeenCalled();
  });
});
