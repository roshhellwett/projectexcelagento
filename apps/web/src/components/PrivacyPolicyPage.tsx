import React from 'react';
import {
  Shield,
  ArrowLeft,
  Lock,
  Building,
  Sparkles,
  ExternalLink,
  CheckCircle2,
} from 'lucide-react';
import { ResourcePageChrome, ResourcePageFooter } from './ResourcePageChrome.js';

interface PrivacyPolicyPageProps {
  onBack: () => void;
  onOpenSettings?: () => void;
  onForgetLearned?: () => void;
  onClearUsage?: () => void;
}

export const PrivacyPolicyPage: React.FC<PrivacyPolicyPageProps> = ({
  onBack,
  onForgetLearned,
  onClearUsage,
}) => {
  return (
    <div className="legal-page resource-page resource-page-privacy" data-testid="privacy-page">
      <ResourcePageChrome current="privacy" onBack={onBack} />
      {/* Top Banner */}
      <div className="legal-top-bar">
        <button className="btn btn-secondary btn-sm" onClick={onBack} title="Return to Workspace">
          <ArrowLeft size={14} /> Back to Workspace
        </button>

        <a
          href="https://zenithopensourceprojects.vercel.app/os"
          target="_blank"
          rel="noopener noreferrer"
          className="zenith-brand-badge"
          title="Visit Zenith Open Source Projects Official Hub"
        >
          <Sparkles size={13} className="zenith-sparkle-icon" />
          <span>Zenith Open Source Projects (zenithopensourceprojects.vercel.app/os)</span>
          <ExternalLink size={12} />
        </a>
      </div>

      <header className="legal-header">
        <div className="legal-badge">
          <Shield size={14} />
          <span>DATA HANDLING · LOCAL AND CONNECTED MODES</span>
        </div>
        <h1 className="legal-title">Privacy Policy &amp; Data Flow Notice</h1>
        <p className="legal-effective">
          This page describes the implemented data flows. It is not an independent security audit or
          legal-compliance certification. Deployments and organizations must verify their own
          provider contracts, access policies, retention settings and legal obligations.
        </p>
        <p className="legal-updated">
          Effective Date: October 2026 • Published by Zenith Open Source Projects
        </p>
      </header>

      {/* Quick Summary Highlights Card */}
      <div className="compliance-highlight-card">
        <div className="highlight-title">
          <Lock size={16} className="text-emerald" />
          <span>Know where your data goes</span>
        </div>
        <div className="highlight-grid">
          <div className="highlight-item">
            <span className="hl-tag">Local spreadsheet engine</span>
            <p>
              Workbook parsing, local calculations and export run in your browser. Parsing/export
              and chat actions/plans use workers where available. Manual edits, proposal previews,
              and deterministic briefings can still run on the main thread. Local mode makes no
              model request. Connected models and optional cloud memory are separate flows described
              below.
            </p>
          </div>
          <div className="highlight-item">
            <span className="hl-tag">Browser-local recovery</span>
            <p>
              The latest successful workbook checkpoint is stored in browser IndexedDB and may
              contain cells, formulas, dates and filename. It is not encrypted by this app. Restore
              or discard it from the workspace; clearing it turns checkpointing off for the session.
            </p>
          </div>
          <div className="highlight-item">
            <span className="hl-tag">Browser-local mission records</span>
            <p>
              The latest 50 tasks are stored in IndexedDB with requests, saved answers, operation
              arguments, previews, evidence, workbook signatures and receipts. These may contain
              customer data and are not encrypted by the app. Remove or clear them from Mission
              control; downloaded JSON contains the same sensitive data. Deleting a mission does not
              delete the workbook checkpoint, learned actions, usage logs or any provider/cloud
              record.
            </p>
          </div>
          <div className="highlight-item">
            <span className="hl-tag">Bring your own key</span>
            <p>
              Inference uses your own API keys (Bring Your Own Key). Keys are stored strictly in
              your browser's origin-scoped storage (`localStorage`) and sent directly over the
              configured connection to your configured LLM provider.
            </p>
          </div>
          <div className="highlight-item">
            <span className="hl-tag">Local data controls</span>
            <p>
              This app provides controls to clear local learned actions, usage logs, mission records
              and workbook checkpoints separately. These controls do not delete provider records or
              optional cloud memory; verify those deletion rights with the relevant provider or
              deployment.
            </p>
          </div>
        </div>
      </div>

      <article className="legal-content">
        <section className="legal-section">
          <h2>1. Introduction &amp; Scope</h2>
          <p>
            This Privacy Policy and Data Protection Notice is issued by{' '}
            <strong>Zenith Open Source Projects</strong> (accessible at{' '}
            <a
              href="https://zenithopensourceprojects.vercel.app/os"
              target="_blank"
              rel="noopener noreferrer"
            >
              https://zenithopensourceprojects.vercel.app/os
            </a>
            ) for the <strong>ExcelAgento</strong> open-source intelligent spreadsheet automation
            platform.
          </p>
          <p>
            The application is distributed as open-source software. Organizations should assess
            their deployment, providers and applicable privacy laws independently; this page is a
            product data-flow description, not a legal-compliance certification.
          </p>
        </section>

        <section className="legal-section">
          <h2>2. Demarcation: Data Fiduciary &amp; Data Processor Roles</h2>
          <p>
            Under Section 2(i) of the DPDP Act 2023, a <em>Data Fiduciary</em> is any person who
            alone or in conjunction with other persons determines the purpose and means of
            processing personal data.
          </p>
          <ul>
            <li>
              <strong>The User is the Data Fiduciary:</strong> When you open, parse, inspect, or
              transform spreadsheets containing personal, financial, or commercial data, you (the
              individual or organization) remain the sole Data Fiduciary. You determine what data is
              loaded and what transformations are performed.
            </li>
            <li>
              <strong>ExcelAgento is a Client-Side Execution Tool:</strong> ExcelAgento is
              distributed as an open-source client-side application. The code executes directly
              within your local browser sandbox (V8/Chromium/WebKit engine) on your hardware. Zenith
              Open Source Projects does not act as a remote data repository or data host.
            </li>
          </ul>
        </section>

        <section className="legal-section">
          <h2>3. Workbook processing, storage and connected context</h2>
          <div className="legal-callout">
            <CheckCircle2 size={16} className="callout-icon text-emerald" />
            <div>
              <strong>Local execution and explicit connected context:</strong> connected models may
              receive your prompt, recent conversation, workbook profiles, examples and requested
              read-tool results when you provide an AI key. Browser-local memory stays separate;
              optional cloud memory is deployment-configured.
            </div>
          </div>
          <p>
            Spreadsheet files (.xlsx, .xls, .csv) use a client-side JavaScript codec. The local
            engine is separate from model inference. Agent memory is browser-local by default;
            optional cloud memory and provider records are separate services with their own
            policies.
          </p>
        </section>

        <section className="legal-section">
          <h2>4. Bring Your Own Key (BYOK) &amp; Artificial Intelligence Model Transit</h2>
          <p>
            ExcelAgento supports optional AI reasoning via Bring Your Own Key (BYOK) providers,
            including Groq, OpenRouter, Google Gemini, and OpenAI:
          </p>
          <ul>
            <li>
              <strong>Local Key Storage:</strong> Your API keys are stored solely in your web
              browser's origin-scoped `localStorage` (`excel_agent_settings_v2`). Any script
              executing on that origin can access this storage; it is not an encrypted secret vault.
              Keys are sent to the configured provider, not a shared-key proxy.
            </li>
            <li>
              <strong>Direct End-to-End Transit:</strong> When an AI request is initiated, your
              browser connects directly over HTTPS to the official endpoint of the provider you
              configured (e.g. <code>api.groq.com</code>, <code>openrouter.ai</code>,{' '}
              <code>generativelanguage.googleapis.com</code>, <code>api.openai.com</code>).
            </li>
            <li>
              <strong>Sanitized Minimal Payloads:</strong> Prior to dispatching schema analysis
              prompts, cell strings are sanitized and truncated. System instructions explicitly mark
              spreadsheet tokens as data to guard against prompt injection.
            </li>
            <li>
              <strong>Deterministic Fallback:</strong> If no API key is provided, ExcelAgento
              executes purely deterministic algorithms locally (local heuristic and formula engines)
              with 0 tokens transmitted over the internet.
            </li>
          </ul>
        </section>

        <section className="legal-section">
          <h2>5. Rights of Data Principals under Indian DPDP Act 2023</h2>
          <p>
            Every citizen and individual whose data is processed has statutory rights guaranteed
            under Chapter III of the Digital Personal Data Protection Act, 2023:
          </p>
          <div className="rights-grid">
            <div className="right-card">
              <h3>Section 11: Right to Access Information</h3>
              <p>
                You have the right to obtain a summary of personal data and processing activities.
                The built-in <strong>Model &amp; Usage Ledger</strong> (`#/usage`) provides an
                unedited, real-time audit record of every request, token count, and destination
                provider.
              </p>
            </div>
            <div className="right-card">
              <h3>Section 12: Right to Correction &amp; Erasure</h3>
              <p>
                You have the right to request immediate correction or erasure of any cached state.
                Clicking the <strong>"Forget Learned Memory"</strong> or{' '}
                <strong>"Clear Usage History"</strong> buttons in the application immediately
                deletes all local patterns and logs.
              </p>
            </div>
            <div className="right-card">
              <h3>Section 13: Right of Grievance Redressal</h3>
              <p>
                You have the right to have grievances redressed by the Data Fiduciary. Zenith Open
                Source Projects provides a designated Grievance Officer whose contact details and
                statutory resolution SLAs are set out below.
              </p>
            </div>
            <div className="right-card">
              <h3>Section 14: Right to Nominate</h3>
              <p>
                You have the right to nominate any other individual to exercise your data protection
                rights in the event of death or incapacity, as provided under the DPDP Act.
              </p>
            </div>
          </div>
        </section>

        <section className="legal-section">
          <h2>6. Support and deployment contacts</h2>
          <p>
            This open-source project does not establish a universal data-controller or retention
            policy for every deployment. For project questions, use the maintainer contact below;
            organizations operating a deployment should publish their own support and privacy
            contact details.
          </p>

          <div className="grievance-officer-box">
            <div className="go-header">
              <Building size={18} className="text-emerald" />
              <span>Project support contact</span>
            </div>
            <div className="go-details">
              <div className="go-detail-row">
                <span className="go-label">Purpose:</span>
                <span className="go-val">Open-source project questions and maintenance</span>
              </div>
              <div className="go-detail-row">
                <span className="go-label">Organization:</span>
                <span className="go-val">Zenith Open Source Projects</span>
              </div>
              <div className="go-detail-row">
                <span className="go-label">Official Portal:</span>
                <span className="go-val">
                  <a
                    href="https://zenithopensourceprojects.vercel.app/os"
                    target="_blank"
                    rel="noopener noreferrer"
                  >
                    https://zenithopensourceprojects.vercel.app/os
                  </a>
                </span>
              </div>
              <div className="go-detail-row">
                <span className="go-label">Official Grievance Email:</span>
                <span className="go-val">
                  <a href="mailto:grievance@zenithopensource.org">grievance@zenithopensource.org</a>{' '}
                  / <a href="mailto:zenithopensource@gmail.com">zenithopensource@gmail.com</a>
                </span>
              </div>
              <div className="go-detail-row">
                <span className="go-label">Deployment policy:</span>
                <span className="go-val">Set by the organization hosting the app</span>
              </div>
            </div>
          </div>
        </section>

        <section className="legal-section">
          <h2>7. Cookies, Web Tracking &amp; Local Storage</h2>
          <p>
            ExcelAgento does <strong>not</strong> employ third-party advertising cookies, cross-site
            trackers, canvas fingerprinting, or profiling telemetry.
          </p>
          <p>
            We utilize standard web browser local storage solely for functional state preservation:
          </p>
          <ul>
            <li>
              <code>excel_agent_settings_v2</code>: Stores your chosen provider, model, and BYOK
              credentials in origin-scoped browser storage; it is not an encrypted secret vault.
            </li>
            <li>
              <code>excel_agent_usage_v1</code>: Local ring-buffer log of token consumption and
              query latency (retained strictly on your machine).
            </li>
            <li>
              <code>excel_agent_theme</code>: User interface visual mode preference (dark/light).
            </li>
            <li>
              <code>excel_agent_memory_v1</code>: Verified learned queries and operation arguments
              stored locally.
            </li>
            <li>
              <code>excelagento-workspace</code>: IndexedDB holding the latest workbook checkpoint;
              it may include cell values, formulas, dates and filename, but not keys, chats or undo
              history.
            </li>
          </ul>
        </section>

        <section className="legal-section">
          <h2>8. Governing Law &amp; Jurisdiction</h2>
          <p>
            This Privacy Policy, and any disputes, actions, claims, or questions relating to data
            protection, shall be governed by, construed, and enforced exclusively in accordance with
            the substantive laws of the <strong>Republic of India</strong>, including the Digital
            Personal Data Protection Act, 2023, the Information Technology Act, 2000, and rules
            promulgated thereunder.
          </p>
          <p>
            Subject to statutory grievance procedures, competent courts located in{' '}
            <strong>Bengaluru, Karnataka, India</strong> or <strong>New Delhi, India</strong> shall
            have exclusive jurisdiction over any proceedings arising out of or in connection with
            this policy.
          </p>
        </section>

        <section className="legal-section">
          <h2>9. Contact &amp; Open Source Governance</h2>
          <p>
            ExcelAgento is built as an open-source flagship of Zenith Open Source Projects. For
            questions or technical audits:
          </p>
          <p>
            Project Repository &amp; Hub:{' '}
            <a
              href="https://zenithopensourceprojects.vercel.app/os"
              target="_blank"
              rel="noopener noreferrer"
            >
              https://zenithopensourceprojects.vercel.app/os
            </a>
            <br />
            Email: <a href="mailto:zenithopensource@gmail.com">zenithopensource@gmail.com</a>
          </p>
        </section>
      </article>

      {/* Action Footer */}
      <footer className="legal-footer">
        <div className="legal-footer-actions">
          {onForgetLearned && (
            <button className="btn btn-secondary btn-sm" onClick={onForgetLearned}>
              Exercise Right to Erasure (Clear Learned Memory)
            </button>
          )}
          {onClearUsage && (
            <button className="btn btn-secondary btn-sm" onClick={onClearUsage}>
              Clear Local Usage Audit Ledger
            </button>
          )}
          <button className="btn btn-primary btn-sm" onClick={onBack}>
            Return to Excel Workspace
          </button>
        </div>
        <p className="legal-copyright">
          &copy; 2026 Zenith Open Source Projects. All Rights Reserved. Licensed under MIT.
        </p>
      </footer>
      <ResourcePageFooter current="privacy" />
    </div>
  );
};
