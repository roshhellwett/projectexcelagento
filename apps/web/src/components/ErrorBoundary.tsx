import React from 'react';

interface ErrorBoundaryProps {
  children: React.ReactNode;
}

interface ErrorBoundaryState {
  error: Error | null;
}

/**
 * Last line of defence: a rendering failure anywhere in the workspace shows a
 * recoverable screen instead of a blank tab.
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

  override render(): React.ReactNode {
    const { error } = this.state;
    if (!error) return this.props.children;

    return (
      <div className="fatal-error">
        <div className="fatal-error-card">
          <div className="fatal-error-title">Something went wrong</div>
          <p className="fatal-error-desc">
            ExcelAgento hit an unexpected error while rendering the workspace. Your data is still
            safe in the browser; reload to continue.
          </p>
          <pre className="fatal-error-detail">{error.message}</pre>
          <button type="button" className="btn btn-primary btn-sm" onClick={this.handleReload}>
            Reload workspace
          </button>
        </div>
      </div>
    );
  }
}
