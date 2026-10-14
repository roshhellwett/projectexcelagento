import React from 'react';
import ReactDOM from 'react-dom/client';
import { App } from './App.js';
import { ErrorBoundary } from './components/ErrorBoundary.js';
import { StartupConfigurationError } from './components/StartupConfigurationError.js';
import { runtimeConfig } from './lib/runtime-config.js';
import './index.css';
import './workbench.css';
import './landing-auth.css';

const rootElement = document.getElementById('root');
if (rootElement) {
  ReactDOM.createRoot(rootElement).render(
    runtimeConfig.issues.length > 0 ? (
      <StartupConfigurationError issues={runtimeConfig.issues} />
    ) : (
      <React.StrictMode>
        <ErrorBoundary>
          <App />
        </ErrorBoundary>
      </React.StrictMode>
    ),
  );
}
