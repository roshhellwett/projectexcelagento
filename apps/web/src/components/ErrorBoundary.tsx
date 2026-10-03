import React from 'react';

interface ErrorBoundaryProps {
  children: React.ReactNode;
  /** When 'panel', render a compact inline card instead of replacing the whole screen. */
  variant?: 'fatal' | 'panel';
  /** Human-readable name of the isolated area, shown in the fallback. */
  label?: string;
}

interface ErrorBoundaryState {
  error: Error | null;
}

/** Stable, user-quotable code derived from the error name and message. */
export function errorCodeFor(error: Error): string {
  const name =
    (error.name || 'Error')
      .replace(/[^A-Za-z]/g, '')
      .slice(0, 3)
      .toUpperCase() || 'ERR';
  let hash = 0;
  const source = `${error.name}:${error.message}`;
  for (let i = 0; i < source.length; i += 1) {
    hash = (hash * 31 + source.charCodeAt(i)) >>> 0;
  }
  return `${name}-${hash.toString(36).toUpperCase().padStart(4, '0').slice(-4)}`;
}

/**
 * Last line of defence: a rendering failure anywhere in the workspace shows a
 * recoverable screen instead of a blank tab. With variant="panel" the failure
 * is isolated to that panel so the rest of the workspace keeps working.
 */
export class ErrorBoundary extends React.Component<ErrorBoundaryProps, ErrorBoundaryState> {
  constructor(props: ErrorBoundaryProps) {
    super(props);
    this.state = { error: null };
  }

  static getDerivedStateFromError(error: Error): ErrorBoundaryState {
    return { error };
  }

  override componentDidCatch(error: Error, info: React.ErrorInfo): void {
    console.error('[ExcelAgento] Unhandled UI error', error, info.componentStack);
  }

  private readonly handleReload = () => {
    window.location.reload();
  };

  private readonly handleDismiss = () => {
    this.setState({ error: null });
  };

  override render(): React.ReactNode {
    const { error } = this.state;
    if (!error) return this.props.children;

    const code = errorCodeFor(error);

    if (this.props.variant === 'panel') {
      return (
        <div className="panel-error" role="alert">
          <div className="panel-error-title">
            Something got broken{this.props.label ? ` in ${this.props.label}` : ''}
          </div>
          <p className="panel-error-desc">
            This panel hit an unexpected error. The rest of your workspace is unaffected. Your data
            is still safe in the browser.
          </p>
          <pre className="fatal-error-detail">
            [{code}] {error.message}
          </pre>
          <button type="button" className="btn btn-secondary btn-sm" onClick={this.handleDismiss}>
            Try again
          </button>
        </div>
      );
    }

    return (
      <div className="fatal-error">
        <div className="fatal-error-card">
          <div className="fatal-error-title">Something got broken</div>
          <p className="fatal-error-desc">
            ExcelAgento hit an unexpected error while rendering the workspace. Your data is still
            safe in the browser; reload to continue.
          </p>
          <pre className="fatal-error-detail">
            [{code}] {error.message}
          </pre>
          <button type="button" className="btn btn-primary btn-sm" onClick={this.handleReload}>
            Reload workspace
          </button>
        </div>
      </div>
    );
  }
}
