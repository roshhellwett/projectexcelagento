import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

describe('deployment authentication connectivity', () => {
  for (const path of ['../../../vercel.json', '../vercel.json']) {
    it(`allows the account API in ${path}`, () => {
      const config = JSON.parse(readFileSync(new URL(path, import.meta.url), 'utf8')) as {
        headers: { headers: { key: string; value: string }[] }[];
      };
      const policy = config.headers
        .flatMap((group) => group.headers)
        .find((header) => header.key === 'Content-Security-Policy')?.value;
      const connections = policy?.split(';').find((rule) => rule.trim().startsWith('connect-src'));
      expect(connections?.split(/\s+/)).toContain('https://fsepapdadtrlddkyqqxu.supabase.co');
    });
  }

  it('loads the theme bootstrap from a CSP-allowed same-origin script', () => {
    const html = readFileSync(new URL('../index.html', import.meta.url), 'utf8');
    expect(html).toContain('<script src="/theme-bootstrap.js"></script>');
    expect(html).not.toContain('<script>\n      // Set the saved or preferred theme');
  });
});
