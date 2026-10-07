import { fileURLToPath } from 'url';
import path from 'path';
import { defineConfig } from 'vitest/config';

const root = path.dirname(fileURLToPath(import.meta.url));

/**
 * Tests resolve workspace packages to their TypeScript sources (mirroring the
 * Vite alias), so `pnpm test` never depends on a prior build artifact.
 */
export default defineConfig({
  resolve: {
    alias: {
      '@excel-agent/engine': path.resolve(root, 'packages/engine/src/index.ts'),
      '@excel-agent/agent': path.resolve(root, 'packages/agent/src/index.ts'),
      '@excel-agent/evals': path.resolve(root, 'packages/evals/src/index.ts'),
    },
  },
  test: {
    environment: 'node',
    globals: true,
    testTimeout: 20000,
    setupFiles: [path.resolve(root, 'apps/web/tests/setup.ts')],
    include: [
      'packages/*/tests/**/*.test.{ts,tsx}',
      'apps/*/tests/**/*.test.{ts,tsx}',
      'apps/*/src/**/*.test.{ts,tsx}',
      'apps/*/src/**/**/*.test.{ts,tsx}',
      'supabase/functions/**/*.test.ts',
    ],
  },
});
