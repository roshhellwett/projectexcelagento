/**
 * Shared Vitest setup.
 *
 * `globals: true` is enabled in vitest.config.ts, which lets Testing Library
 * register its own automatic `afterEach(cleanup)` hook. This file only extends
 * the matchers, so it stays harmless for the node-environment package tests.
 */
import '@testing-library/jest-dom/vitest';
