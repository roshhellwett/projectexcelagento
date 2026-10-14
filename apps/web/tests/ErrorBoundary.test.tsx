// @vitest-environment jsdom
import { describe, expect, it, vi } from 'vitest';
import { render, screen } from '@testing-library/react';

import { ErrorBoundary } from '../src/components/ErrorBoundary.js';

const Boom = () => {
  throw new Error('kaboom');
};

describe('ErrorBoundary', () => {
  it('renders children when nothing throws', () => {
    render(
      <ErrorBoundary>
        <p>safe content</p>
      </ErrorBoundary>,
    );

    expect(screen.getByText('safe content')).toBeInTheDocument();
  });

  it('shows a recoverable, privacy-safe screen instead of a blank tab', () => {
    // React reports boundary-caught errors to the console; keep test output clean.
    const consoleError = vi.spyOn(console, 'error').mockImplementation(() => {});

    render(
      <ErrorBoundary>
        <Boom />
      </ErrorBoundary>,
    );

    expect(screen.getByText('Something got broken')).toBeInTheDocument();
    expect(screen.getByText(/Reference code:/)).toBeInTheDocument();
    expect(screen.queryByText(/kaboom/)).not.toBeInTheDocument();
    expect(screen.getByRole('button', { name: /Reload workspace/i })).toBeInTheDocument();
    expect(screen.queryByText('safe content')).not.toBeInTheDocument();
    const boundaryLogs = consoleError.mock.calls.filter(
      ([label]) => label === '[ExcelAgento] Unhandled UI error',
    );
    expect(boundaryLogs).toHaveLength(1);
    expect(JSON.stringify(boundaryLogs)).not.toContain('kaboom');

    consoleError.mockRestore();
  });
});
