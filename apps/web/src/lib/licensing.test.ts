import { describe, expect, it } from 'vitest';

import { resolvePlaintextLicenseKey } from './licensing.js';

describe('resolvePlaintextLicenseKey', () => {
  it('uses the stored plaintext only when it exists', () => {
    expect(
      resolvePlaintextLicenseKey(
        { raw_key: 'EXCEL-AAAAAAAA-BBBBBBBB-CCCCCCCC-DDDDDDDD', key_hint: 'EXCEL-••••-DDDDDD' },
        [],
      ),
    ).toBe('EXCEL-AAAAAAAA-BBBBBBBB-CCCCCCCC-DDDDDDDD');
  });

  it('matches a newly generated key in the current admin session', () => {
    expect(
      resolvePlaintextLicenseKey({ raw_key: null, key_hint: 'EXCEL-••••-8D378A' }, [
        'EXCEL-01234567-89ABCDEF-01234567-8D378A',
      ]),
    ).toBe('EXCEL-01234567-89ABCDEF-01234567-8D378A');
  });

  it('never returns a masked hint as a usable key', () => {
    expect(
      resolvePlaintextLicenseKey({ raw_key: null, key_hint: 'EXCEL-••••-8D378A' }, []),
    ).toBeUndefined();
  });
});
