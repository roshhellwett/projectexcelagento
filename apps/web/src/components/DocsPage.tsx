import React, { useState } from 'react';
import {
  BookOpen,
  ArrowLeft,
  Sparkles,
  ExternalLink,
  Cpu,
  ShieldCheck,
  Calculator,
  Database,
} from 'lucide-react';

interface DocsPageProps {
  onBack: () => void;
  onOpenSettings?: () => void;
}

export const DocsPage: React.FC<DocsPageProps> = ({ onBack }) => {
  const [activeTab, setActiveTab] = useState<
    'architecture' | 'quickstart' | 'operations' | 'agents' | 'privacy'
  >('architecture');

  return (
    <div className="docs-page" data-testid="docs-page">
      {/* Top Bar with Zenith OS Branding */}
      <div className="docs-top-bar">
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

      {/* Docs Header */}
      <header className="docs-header">
        <div className="docs-badge">
          <BookOpen size={14} />
          <span>OFFICIAL DEVELOPER &amp; ARCHITECTURE MANUAL</span>
        </div>
        <h1 className="docs-title">
          ExcelAgento <span className="text-emerald">Documentation</span>
        </h1>
        <p className="docs-subtitle">
          The open-source spreadsheet workspace by Zenith Open Source Projects. It combines
          browser-local workbook operations with an optional bring-your-own-key reasoning layer,
          reviewable previews, and reversible history.
        </p>

        {/* Tab Navigation */}
        <div className="docs-tabs">
          <button
            className={`docs-tab-btn ${activeTab === 'architecture' ? 'active' : ''}`}
            onClick={() => setActiveTab('architecture')}
          >
            System Architecture
          </button>
          <button
            className={`docs-tab-btn ${activeTab === 'quickstart' ? 'active' : ''}`}
            onClick={() => setActiveTab('quickstart')}
          >
            Quickstart &amp; BYOK
          </button>
          <button
            className={`docs-tab-btn ${activeTab === 'operations' ? 'active' : ''}`}
            onClick={() => setActiveTab('operations')}
          >
            Engine Operations
          </button>
          <button
            className={`docs-tab-btn ${activeTab === 'agents' ? 'active' : ''}`}
            onClick={() => setActiveTab('agents')}
          >
            Agent Workforce
          </button>
          <button
            className={`docs-tab-btn ${activeTab === 'privacy' ? 'active' : ''}`}
            onClick={() => setActiveTab('privacy')}
          >
            DPDP Act 2023 Compliance
          </button>
        </div>
      </header>

      {/* Content Area */}
      <main className="docs-content">
        {activeTab === 'architecture' && (
          <section className="docs-section">
            <h2>1. System Architecture Overview</h2>
            <p>
              ExcelAgento is built on a <strong>layered multi-agent architecture</strong> designed
              to bridge the flexibility of LLM reasoning with the strict reliability of
              transactional spreadsheet software.
            </p>

            <div className="docs-card-grid">
              <div className="docs-card">
                <div className="docs-card-header">
                  <Cpu className="text-emerald" size={20} />
                  <h3>Browser-local workbook runtime</h3>
                </div>
                <p>
                  Spreadsheet files (.xlsx, .xls, .csv) are parsed in the browser. Workbook
                  transformations, deterministic analysis, checkpoints, and history stay in the
                  local app. When you connect a provider, the request and bounded context may be
                  sent directly to that provider.
                </p>
              </div>

              <div className="docs-card">
                <div className="docs-card-header">
                  <Calculator className="text-blue" size={20} />
                  <h3>Deterministic Math Core</h3>
                </div>
                <p>
                  Supported calculations and statistics are computed by the engine rather than raw
                  LLM token completion. The workspace reports exclusions, unsupported formulas, and
                  undefined results instead of inventing values.
                </p>
              </div>

              <div className="docs-card">
                <div className="docs-card-header">
                  <ShieldCheck className="text-pink" size={20} />
                  <h3>Pre-Flight Invariant Guardrails</h3>
                </div>
                <p>
                  Every proposed operation passes through the <strong>Aegis Guardrail</strong>{' '}
                  layer. Operations that affect 0 cells or could wipe data are blocked or require
                  explicit human confirmation with a cell-diff preview.
                </p>
              </div>

              <div className="docs-card">
                <div className="docs-card-header">
                  <Database className="text-purple" size={20} />
                  <h3>In-Browser Local Neural Memory</h3>
                </div>
                <p>
                  Verified action associations and working session context are kept in browser-local
                  storage. Provider requests remain separate from local memory and are visible in
                  the Model &amp; Usage ledger when a key is configured.
                </p>
              </div>
            </div>

            <div className="docs-code-block">
              <div className="code-header">Agent Dispatch Pipeline</div>
              <pre>
                {`[User Input / Query]
       │
       ▼
[Typo-Tolerant NLP Normalizer] ─── Canonicalizes queries (e.g. "remov dupli" → "dedup")
       │
       ▼
[Nexus Conductor] ────────────── Analyzes Intent & Decomposes Multi-Step Plans
       │
  ┌────┴───────────────────────────┐
  ▼                                ▼
[Atlas Quantitative Calc]    [Valkyrie Worker Thread]
(CAGR, Margins, Stats)        (Dedup, Dates, Pivots)
  │                                │
  └────┬───────────────────────────┘
       ▼
[Aegis Guardrail Layer] ──────── Validates bounds, dry-run diff preview
       │
       ▼
[Transactional Workbook State] ── Immutable history stack (Undo / Redo)
       │
       ▼
[Cortex Memory Engine] ───────── Caches successful patterns with zero cell leakage`}
              </pre>
            </div>
          </section>
        )}

        {activeTab === 'quickstart' && (
          <section className="docs-section">
            <h2>2. Quickstart &amp; Bring Your Own Key (BYOK)</h2>
            <p>
              ExcelAgento can be used immediately with built-in deterministic operations, or powered
              by state-of-the-art LLMs using your own API key.
            </p>

            <div className="docs-steps-list">
              <div className="docs-step-item">
                <div className="step-badge">Step 1</div>
                <div className="step-body">
                  <h3>Load Your Workbook</h3>
                  <p>
                    Click <strong>Upload File</strong> in the top navigation bar to load any Excel
                    (.xlsx, .xls) or CSV file. You can also select pre-built sample fixtures from
                    the dropdown menu to test messy data formats.
                  </p>
                </div>
              </div>

              <div className="docs-step-item">
                <div className="step-badge">Step 2</div>
                <div className="step-body">
                  <h3>Configure Your AI Model (Optional)</h3>
                  <p>
                    Click the <strong>Settings</strong> icon (or click "Change Key" in the copilot
                    panel) to enter your API key. ExcelAgento natively supports:
                  </p>
                  <ul>
                    <li>
                      <strong>OpenRouter:</strong> Access to all open models (e.g., Llama 3,
                      DeepSeek, Claude, Mistral)
                    </li>
                    <li>
                      <strong>Groq:</strong> Ultra-low-latency 500+ tok/s inference
                    </li>
                    <li>
                      <strong>Google Gemini:</strong> Gemini 2.5 Flash / Pro models
                    </li>
                    <li>
                      <strong>OpenAI:</strong> GPT-4o / GPT-4o-mini
                    </li>
                  </ul>
                  <p className="text-muted">
                    Your key is stored solely in your browser's local sandbox and is never
                    transmitted to Zenith Open Source Projects.
                  </p>
                </div>
              </div>

              <div className="docs-step-item">
                <div className="step-badge">Step 3</div>
                <div className="step-body">
                  <h3>Instruct the Agents in Plain English</h3>
                  <p>
                    Type any command into the copilot chat, or use keyboard shortcut{' '}
                    <code>Cmd+K</code> / <code>Ctrl+K</code> to summon the Command Palette.
                  </p>
                </div>
              </div>
            </div>
          </section>
        )}

        {activeTab === 'operations' && (
          <section className="docs-section">
            <h2>3. Supported Engine Operations</h2>
            <p>ExcelAgento features a rich library of 16+ transactional spreadsheet operations:</p>

            <div className="docs-ops-table-wrapper">
              <table className="docs-ops-table">
                <thead>
                  <tr>
                    <th>Operation</th>
                    <th>Category</th>
                    <th>Description</th>
                    <th>Sample Prompt</th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td>
                      <code>remove_duplicates</code>
                    </td>
                    <td>Clean</td>
                    <td>Deduplicates rows by single or composite columns</td>
                    <td>"Remove duplicate order IDs"</td>
                  </tr>
                  <tr>
                    <td>
                      <code>standardize_dates</code>
                    </td>
                    <td>Format</td>
                    <td>Converts mixed UK, US, serial, and slash dates to ISO 8601</td>
                    <td>"Format dates in column C to YYYY-MM-DD"</td>
                  </tr>
                  <tr>
                    <td>
                      <code>sort_range</code>
                    </td>
                    <td>Transform</td>
                    <td>Sorts alphanumeric and numeric columns ascending or descending</td>
                    <td>"Sort rows by Revenue descending"</td>
                  </tr>
                  <tr>
                    <td>
                      <code>filter_rows</code>
                    </td>
                    <td>Filter</td>
                    <td>Extracts matching rows into a new dedicated sheet</td>
                    <td>"Filter rows where status is Active"</td>
                  </tr>
                  <tr>
                    <td>
                      <code>aggregate_column</code>
                    </td>
                    <td>Analytics</td>
                    <td>Computes Sum, Average, Min, Max, Count, or Median</td>
                    <td>"Calculate total and average of Amount"</td>
                  </tr>
                  <tr>
                    <td>
                      <code>calculate_growth_rate</code>
                    </td>
                    <td>Analytics</td>
                    <td>Computes multi-period CAGR or annual growth rate</td>
                    <td>"Calculate 10-year CAGR for Revenue"</td>
                  </tr>
                  <tr>
                    <td>
                      <code>text_to_numbers</code>
                    </td>
                    <td>Clean</td>
                    <td>Strips currency symbols and parses numbers stored as text</td>
                    <td>"Convert column D numbers stored as text"</td>
                  </tr>
                  <tr>
                    <td>
                      <code>trim_whitespace</code>
                    </td>
                    <td>Clean</td>
                    <td>Cleans leading, trailing, and double interior spaces</td>
                    <td>"Trim spaces across all text columns"</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </section>
        )}

        {activeTab === 'agents' && (
          <section className="docs-section">
            <h2>4. The Autonomous Multi-Agent Workforce</h2>
            <p>
              Learn about the six specialized neural agents in ExcelAgento. For interactive profiles
              and command testing, visit the dedicated <a href="#/agents">Agents Page</a>.
            </p>
            <ul>
              <li>
                <strong>Nexus Conductor:</strong> Intent classification, plan decomposition,
                proactive clarification.
              </li>
              <li>
                <strong>Atlas Data Scientist:</strong> 64-bit deterministic calculations, CAGR, tax
                rates, margins.
              </li>
              <li>
                <strong>Aegis Guardrail:</strong> Schema validation, destructive action gatekeeping,
                diff previews.
              </li>
              <li>
                <strong>Cortex Memory Engine:</strong> Browser-local verified action memory and
                working session context.
              </li>
              <li>
                <strong>Valkyrie Transformer:</strong> Worker-assisted parsing and export where
                available, with local engine transformations and a fallback path.
              </li>
              <li>
                <strong>Scout Knowledge Agent:</strong> Autonomous formula synthesis (XLOOKUP,
                INDEX/MATCH).
              </li>
            </ul>
          </section>
        )}

        {activeTab === 'privacy' && (
          <section className="docs-section">
            <h2>5. Indian DPDP Act 2023 Statutory Compliance</h2>
            <p>
              ExcelAgento is fully compliant with the{' '}
              <strong>Digital Personal Data Protection Act, 2023</strong>:
            </p>
            <ul>
              <li>
                <strong>Zero Data Ingestion:</strong> Your spreadsheet contents are never uploaded
                to remote servers.
              </li>
              <li>
                <strong>Data Principal Rights:</strong> Exercise your Right to Erasure anytime by
                clearing learned memory.
              </li>
              <li>
                <strong>Statutory Grievance Redressal:</strong> Grievance Officer stationed in
                Bengaluru, Karnataka, India with 15-day statutory resolution guarantee.
              </li>
            </ul>
            <p>
              For full legal details, see our <a href="#/privacy">Privacy Policy</a>.
            </p>
          </section>
        )}
      </main>

      <footer className="docs-footer">
        <a
          href="https://zenithopensourceprojects.vercel.app/os"
          target="_blank"
          rel="noopener noreferrer"
          className="text-emerald"
        >
          Zenith Open Source Projects Hub (https://zenithopensourceprojects.vercel.app/os)
        </a>
        <span>&bull;</span>
        <a href="#/agents">Agents Showcase</a>
        <span>&bull;</span>
        <a href="#/privacy">Privacy Policy (DPDP 2023)</a>
        <span>&bull;</span>
        <a href="#/terms">Terms of Service</a>
      </footer>
    </div>
  );
};
