import { describe, expect, it } from 'vitest';
import { resolveSupabaseClientConfig, validateRuntimeConfig } from './runtime-config.js';

const productionEnvironment = {
  MODE: 'production',
  VITE_SUPABASE_URL: 'https://project.example.supabase.co',
  VITE_SUPABASE_ANON_KEY: 'sb_publishable_synthetic_anon_key_1234',
  VITE_AUTH_REDIRECT_ORIGIN: 'https://excelagento.example.com',
};

const legacyAnonKey =
  'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJyb2xlIjoiYW5vbiJ9.synthetic-signature';

describe('production runtime configuration', () => {
  it('accepts explicit HTTPS auth configuration without exposing its values in issues', () => {
    const config = validateRuntimeConfig(productionEnvironment);

    expect(config.issues).toEqual([]);
    expect(config.supabaseUrlValid).toBe(true);
    expect(config.supabaseAnonKeyValid).toBe(true);
    expect(config.authRedirectOrigin).toBe('https://excelagento.example.com');
  });

  it('rejects incomplete or unsafe production values by variable name only', () => {
    const privateValue = 'do-not-echo-this-value';
    const config = validateRuntimeConfig({
      MODE: 'production',
      VITE_SUPABASE_URL: 'http://localhost:54321?secret=do-not-echo-this-value',
      VITE_SUPABASE_ANON_KEY: `sb_secret_${privateValue}`,
      VITE_AUTH_REDIRECT_ORIGIN: 'http://localhost:5173',
    });

    expect(config.issues).toEqual([
      'VITE_SUPABASE_URL must be a valid HTTPS URL without embedded credentials.',
      'VITE_SUPABASE_ANON_KEY must be a public Supabase publishable key or a legacy JWT anon key; private or placeholder keys are not allowed.',
      'VITE_AUTH_REDIRECT_ORIGIN must be a valid HTTPS origin.',
    ]);
    expect(config.issues.join(' ')).not.toContain(privateValue);
  });

  it.each([
    'sb_secret_0123456789abcdef0123456789',
    'eyJhbGciOiJIUzI1NiJ9.eyJyb2xlIjoic2VydmljZV9yb2xlIn0.synthetic-signature',
    'a placeholder value',
  ])('rejects private or placeholder key shapes without echoing the key: %s', (key) => {
    const config = validateRuntimeConfig({
      ...productionEnvironment,
      VITE_SUPABASE_ANON_KEY: key,
    });

    expect(config.supabaseAnonKeyValid).toBe(false);
    expect(config.issues).toContain(
      'VITE_SUPABASE_ANON_KEY must be a public Supabase publishable key or a legacy JWT anon key; private or placeholder keys are not allowed.',
    );
    expect(config.issues.join(' ')).not.toContain(key);
  });

  it('accepts both publishable and legacy JWT anon keys', () => {
    const publishable = validateRuntimeConfig(productionEnvironment);
    const legacy = validateRuntimeConfig({
      ...productionEnvironment,
      VITE_SUPABASE_ANON_KEY: legacyAnonKey,
    });

    expect(publishable.supabaseAnonKeyValid).toBe(true);
    expect(publishable.issues).toEqual([]);
    expect(legacy.supabaseAnonKeyValid).toBe(true);
    expect(legacy.issues).toEqual([]);
  });

  it('rejects unsupported publishable and legacy JWT placeholder shapes', () => {
    const publishablePlaceholder = validateRuntimeConfig({
      ...productionEnvironment,
      VITE_SUPABASE_ANON_KEY: 'sb_publishable_placeholder_key',
    });
    const malformedLegacy = validateRuntimeConfig({
      ...productionEnvironment,
      VITE_SUPABASE_ANON_KEY: 'eyJhbGciOiJIUzI1NiJ9.not-json.synthetic-signature',
    });

    expect(publishablePlaceholder.supabaseAnonKeyValid).toBe(false);
    expect(malformedLegacy.supabaseAnonKeyValid).toBe(false);
  });

  it('keeps control-character diagnostics redacted', () => {
    const secret = 'sb_secret_private-value';
    const config = validateRuntimeConfig({
      ...productionEnvironment,
      VITE_SUPABASE_ANON_KEY: `${secret}\nvalue`,
    });

    expect(config.issues).toContain('VITE_SUPABASE_ANON_KEY contains invalid control characters.');
    expect(config.issues.join(' ')).not.toContain(secret);
  });

  it('keeps development fallbacks available while production clients fail closed', () => {
    const development = validateRuntimeConfig({ MODE: 'development' });
    expect(development.issues).toEqual([]);

    const production = validateRuntimeConfig({ MODE: 'production' });
    const client = resolveSupabaseClientConfig(production, {
      url: 'https://checked-in-default.example',
      anonKey: 'checked-in-default-key',
    });

    expect(client.disabled).toBe(true);
    expect(client.url).toBe('https://invalid.excelagento.invalid');
    expect(client.anonKey).toBe('missing-production-configuration');
  });

  it('fails closed for custom deployment modes instead of using checked-in defaults', () => {
    const staging = validateRuntimeConfig({ MODE: 'staging' });
    expect(staging.isProduction).toBe(false);
    expect(staging.issues).toEqual([
      'VITE_SUPABASE_URL is required for production authentication.',
      'VITE_SUPABASE_ANON_KEY is required for production authentication.',
      'VITE_AUTH_REDIRECT_ORIGIN is required for production email links.',
    ]);

    const client = resolveSupabaseClientConfig(staging, {
      url: 'https://checked-in-default.example',
      anonKey: 'checked-in-default-key',
    });
    expect(client.disabled).toBe(true);
    expect(client.url).toBe('https://invalid.excelagento.invalid');
  });

  it('rejects paths and query strings in deployment origins', () => {
    const config = validateRuntimeConfig({
      MODE: 'production',
      VITE_SUPABASE_URL: 'https://project.example.supabase.co/project',
      VITE_SUPABASE_ANON_KEY: productionEnvironment.VITE_SUPABASE_ANON_KEY,
      VITE_AUTH_REDIRECT_ORIGIN: 'https://excelagento.example.com/login?source=email',
    });

    expect(config.issues).toEqual([
      'VITE_SUPABASE_URL must be a valid HTTPS URL without embedded credentials.',
      'VITE_AUTH_REDIRECT_ORIGIN must be a valid HTTPS origin.',
    ]);
  });
});
