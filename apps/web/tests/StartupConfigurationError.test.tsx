// @vitest-environment jsdom
import { describe, expect, it, vi } from 'vitest';
import { render, screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';

import { StartupConfigurationError } from '../src/components/StartupConfigurationError.js';

describe('StartupConfigurationError', () => {
  it('explains the blocked startup without rendering secret values', async () => {
    const reload = vi.fn();
    const user = userEvent.setup();

    render(
      <StartupConfigurationError
        issues={['VITE_SUPABASE_URL is required for production authentication.']}
        onReload={reload}
      />,
    );

    expect(screen.getByRole('alert')).toHaveTextContent('Deployment configuration needed');
    expect(screen.getByText(/VITE_SUPABASE_URL is required/)).toBeInTheDocument();
    expect(screen.getByText(/No workbook or provider data was loaded/)).toBeInTheDocument();

    await user.click(screen.getByRole('button', { name: /Reload after fixing configuration/i }));
    expect(reload).toHaveBeenCalledTimes(1);
  });
});
