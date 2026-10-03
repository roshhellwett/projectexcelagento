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

  it('shows a recoverable screen instead of a blank tab', () => {
    // React reports boundary-caught errors to the console; keep test output clean.
    const consoleError = vi.spyOn(console, 'error').mockImplementation(() => {});

    render(
      <ErrorBoundary>
        <Boom />
      </ErrorBoundary>,
    );

    expect(screen.getByText('Something went wrong')).toBeInTheDocument();
    expect(screen.getByText(/kaboom/)).toBeInTheDocument();
    expect(screen.getByRole('button', { name: /Reload workspace/i })).toBeInTheDocument();
    expect(screen.queryByText('safe content')).not.toBeInTheDocument();

    consoleError.mockRestore();
  });
});
