// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { authEmailRedirectUrl, readAuthCallback } from '../src/lib/auth-redirect.js';

describe('email destinations and callback recognition', () => {
  it('uses the deployed origin and base path, excluding unrelated query and routing fragments', () => {
    expect(
      authEmailRedirectUrl('confirm', 'https://excelagento.example/app/?campaign=signup#/auth'),
    ).toBe('https://excelagento.example/app/?auth=confirm');
    expect(authEmailRedirectUrl('recovery', 'http://localhost:5173/#/auth')).toBe(
      'http://localhost:5173/?auth=recovery',
    );
  });

  it('recognizes confirmation and recovery callbacks after the SDK has consumed their fragments', () => {
    expect(readAuthCallback('https://excelagento.example/?auth=confirm')).toEqual({
      kind: 'confirm',
      error: null,
    });
    expect(readAuthCallback('https://excelagento.example/?auth=recovery')).toEqual({
      kind: 'recovery',
      error: null,
    });
    expect(
      readAuthCallback('https://excelagento.example/#access_token=test&type=recovery'),
    ).toEqual({ kind: 'recovery', error: null });
  });

  it('recognizes expired links while treating ordinary hash routes as navigation', () => {
    expect(readAuthCallback('https://excelagento.example/#/signup')).toBeNull();
    expect(
      readAuthCallback(
        'https://excelagento.example/?auth=confirm#error=access_denied&error_code=otp_expired',
      )?.error,
    ).toContain('expired or has already been used');
  });
});
