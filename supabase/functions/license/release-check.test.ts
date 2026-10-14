import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { inspectLicenseEntry } from '../../../tools/release-check-rules.mjs';

describe('license release hygiene checks', () => {
  it('detects non-empty pepper fallbacks with the actual Deno environment syntax', () => {
    const errors = inspectLicenseEntry(
      `
        const keyPepper = Deno.env.get('LICENSE_KEY_PEPPER') ||
          'checked-in-pepper-that-must-not-be-used';
      `,
    );

    expect(errors).toContain(
      'license Edge Function must not fall back to a checked-in key pepper.',
    );
  });

  it('rejects every non-empty direct APP_ORIGIN fallback, including wildcard origins', () => {
    expect(
      inspectLicenseEntry("Deno.env.get('APP_ORIGIN') ?? 'https://checked-in.example';"),
    ).toContain('license Edge Function must not use a non-empty APP_ORIGIN fallback literal.');
    expect(inspectLicenseEntry("Deno.env.get('APP_ORIGIN') || '*';")).toContain(
      'license Edge Function must not default to a wildcard application origin.',
    );
  });

  it('allows empty fail-closed fallbacks and the current license entry point', () => {
    expect(
      inspectLicenseEntry(
        `
          const origin = Deno.env.get('APP_ORIGIN') ?? '';
          const pepper = Deno.env.get('LICENSE_KEY_PEPPER') || '';
        `,
      ),
    ).toEqual([]);

    const currentEntry = readFileSync(new URL('./index.ts', import.meta.url), 'utf8');
    expect(inspectLicenseEntry(currentEntry)).toEqual([]);
  });

  it('keeps generated plaintext keys out of the database after the hardening migration', () => {
    const migration = readFileSync(
      new URL('../../migrations/20261013000000_harden_activation_key_storage.sql', import.meta.url),
      'utf8',
    );

    expect(migration).toContain('UPDATE public.license_keys');
    expect(migration).toContain('SET raw_key = NULL');
    expect(migration).toContain('NEW.raw_key := NULL');
    expect(migration).toContain('CREATE TRIGGER license_keys_clear_plaintext_key');
  });
});
