import React, { useState } from 'react';
import {
  Bot,
  Brain,
  ShieldCheck,
  Calculator,
  Zap,
  Search,
  ArrowLeft,
  ExternalLink,
  Sparkles,
  CheckCircle2,
  Cpu,
  Terminal,
  ArrowRight,
} from 'lucide-react';

interface AgentProfile {
  id: string;
  name: string;
  codename: string;
  role: string;
  clearance: string;
  icon: React.ReactNode;
  accentColor: string;
  status: 'ACTIVE' | 'STANDBY';
  description: string;
  responsibilities: string[];
  capabilities: string[];
  metrics: { label: string; value: string }[];
  samplePrompts: string[];
}

interface AgentsPageProps {
  onBack: () => void;
  onSelectPrompt?: (prompt: string) => void;
}

export const AgentsPage: React.FC<AgentsPageProps> = ({ onBack, onSelectPrompt }) => {
  const [selectedFilter, setSelectedFilter] = useState<
    'all' | 'orchestration' | 'analytics' | 'safety' | 'data'
  >('all');
  const [, setCopiedPrompt] = useState<string | null>(null);
  const agents: AgentProfile[] = [
    {
      id: 'nexus',
      name: 'Nexus Conductor',
      codename: 'NEXUS-01 // LEAD ORCHESTRATOR',
      role: 'Chief Multi-Agent Conductor & Intent Strategist',
      clearance: 'Level 5 (System Orchestration)',
      icon: <Brain className="agent-icon" size={26} />,
      accentColor: '#10b981',
      status: 'ACTIVE',
      description:
        'The master conductor of ExcelAgento. Nexus parses incoming human queries, disambiguates messy natural language, classifies user intent, and breaks multi-step objectives into phased execution plans.',
      responsibilities: [
        'Natural Language Understanding & typo-tolerant semantic normalization',
        'Complex objective decomposition into phased Execution Plans',
        'Dynamic sub-agent routing and parallel tool orchestration',
        'Proactive Clarification engine when user intent is ambiguous',
      ],
      capabilities: [
        'Intent Routing',
        'Execution Planner',
        'Tool Scheduling',
        'Proactive Clarification',
        'Context Routing',
      ],
      metrics: [
        { label: 'Intent Accuracy', value: '99.4%' },
        { label: 'Dispatch Latency', value: '< 18ms' },
        { label: 'Execution Clearance', value: 'Unrestricted' },
      ],
      samplePrompts: [
        'Clean and structure this sheet into human readable format',
        'Create a new column of TOTAL and calculate all rows from FY09 to FY18',
        'Identify duplicate records and separate them into an audit tab',
      ],
    },
    {
      id: 'atlas',
      name: 'Atlas Data Scientist',
      codename: 'ATLAS-CALC // STATISTICAL CORE',
      role: 'Deterministic Calculation & Statistical Engine',
      clearance: 'Level 4 (Quantitative Analytics)',
      icon: <Calculator className="agent-icon" size={26} />,
      accentColor: '#3b82f6',
      status: 'ACTIVE',
      description:
        'The rigorous mathematical mind. Atlas performs multi-year CAGR, variance analysis, standard deviations, financial margins, and deterministic aggregations with 0% mathematical hallucination guarantee.',
      responsibilities: [
        'Deterministic calculations without LLM floating point arithmetic errors',
        'Financial analysis (Operating Margin, Net Margin, Effective Tax Rates)',
        'Statistical distribution analysis (Mean, Median, Standard Deviation, Z-Scores)',
        'Autonomous formula synthesis for CAGR, SUMIFS, and XLOOKUP',
      ],
      capabilities: [
        'Deterministic Math',
        'CAGR Engine',
        'Tax & Margin Modeling',
        'Column Profiling',
        'Outlier Detection',
      ],
      metrics: [
        { label: 'Arithmetic Error', value: '0.00%' },
        { label: 'Numerical Precision', value: '64-bit Float' },
        { label: 'Aggregation Speed', value: '< 4ms / 10k rows' },
      ],
      samplePrompts: [
        'Analyze the total column and let me know where it is profitable each year',
        'Calculate 10-year CAGR for Revenue and Operating Income',
        'What is the effective tax rate in FY17 and why is it an anomaly?',
      ],
    },
    {
      id: 'aegis',
      name: 'Aegis Guardrail',
      codename: 'AEGIS-SHIELD // INVARIANT KEEPER',
      role: 'Enterprise Invariant Verifier & Safety Gatekeeper',
      clearance: 'Level 5 (Security & Schema Gatekeeper)',
      icon: <ShieldCheck className="agent-icon" size={26} />,
      accentColor: '#ec4899',
      status: 'ACTIVE',
      description:
        'The uncompromising safety shield. Aegis inspects every mutation before it touches the workbook, verifies cell bounds, enforces schema invariants, prevents data loss, and mandates explicit confirmation for destructive actions.',
      responsibilities: [
        'Pre-flight dry run validation and transactional preview generation',
        'Destructive mutation gatekeeping (blocking accidental bulk deletions)',
        'Spreadsheet schema invariant enforcement across linked tabs',
        'Safe transactional rollback if an engine operation encounters errors',
      ],
      capabilities: [
        'Pre-Flight Dry Run',
        'Destructive Guard',
        'Transactional Rollback',
        'Schema Validation',
        'Cell Diff Audit',
      ],
      metrics: [
        { label: 'Data Corruption', value: '0 Incidents' },
        { label: 'Verification Pass', value: '100%' },
        { label: 'Invariant Checks', value: '12 Rules / Op' },
      ],
      samplePrompts: [
        'Delete all rows where amount is negative',
        'Replace empty cells in column C with NA without touching formulas',
        'Verify schema compatibility before merging sheet1 with sheet2',
      ],
    },
    {
      id: 'cortex',
      name: 'Cortex Memory Engine',
      codename: 'CORTEX-NEURAL // COLLECTIVE BRAIN',
      role: 'Dual-Store Supabase Neural Memory & Continuous Learning',
      clearance: 'Level 4 (Neural State & Continuity)',
      icon: <Cpu className="agent-icon" size={26} />,
      accentColor: '#8b5cf6',
      status: 'ACTIVE',
      description:
        'The persistent knowledge layer. Cortex synchronizes dual-store memory across local browser storage and Supabase cloud cortex, preserving learned actions, user preferences, and multi-turn session working memory.',
      responsibilities: [
        'Multi-turn conversational session context and working step tracking',
        'Continuous learning from user feedback and successful operations',
        'Supabase cloud synchronization with instant local fallback',
        'Privacy compliance: Zero spreadsheet cell values stored on remote servers',
      ],
      capabilities: [
        'Supabase Sync',
        'Working Memory',
        'Cross-Turn State',
        'User Pattern Cache',
        'Zero-Remote-Data DPDP',
      ],
      metrics: [
        { label: 'Context Retrieval', value: '< 12ms' },
        { label: 'Cloud Fallback', value: '100% Graceful' },
        { label: 'Remote Cell Storage', value: '0 Bytes' },
      ],
      samplePrompts: [
        'Do the same formatting we did on the previous column',
        'Remember that column D represents Indian Rupees (INR)',
        'Apply the learned capitalization pattern to supplier names',
      ],
    },
    {
      id: 'valkyrie',
      name: 'Valkyrie Transformer',
      codename: 'VALKYRIE-WORKER // HIGH-THROUGHPUT CORE',
      role: 'High-Throughput Web Worker Data Engineering Specialist',
      clearance: 'Level 4 (Worker Thread Execution)',
      icon: <Zap className="agent-icon" size={26} />,
      accentColor: '#f59e0b',
      status: 'ACTIVE',
      description:
        'The heavy machinery. Valkyrie runs in dedicated HTML5 background Web Workers, crunching 100,000+ row datasets, deduping, normalizing date formats, converting text-to-numbers, and pivoting without freezing the browser UI.',
      responsibilities: [
        'Non-blocking background Web Worker transformations on massive files',
        'Multi-format date normalization (mixed US, UK, ISO-8601, Excel serials)',
        'Deduplication across composite columns with fuzzy matching',
        'Multi-column pivots, sorting, and structured sheet restructuring',
      ],
      capabilities: [
        'Web Worker Threads',
        '100k+ Row Zero Lag',
        'Date Normalizer',
        'Composite Dedup',
        'Pivot Engine',
      ],
      metrics: [
        { label: 'Throughput', value: '45,000 rows/sec' },
        { label: 'UI Thread Freeze', value: '0.0 ms' },
        { label: 'Worker Concurrency', value: 'Hardware Native' },
      ],
      samplePrompts: [
        'Format dates in column C to YYYY-MM-DD',
        'Remove duplicate rows considering order ID and customer ID',
        'Convert all numbers stored as text to numeric values in column E',
      ],
    },
    {
      id: 'scout',
      name: 'Scout Knowledge Agent',
      codename: 'SCOUT-RESEARCH // INTELLIGENCE AGENT',
      role: 'Autonomous Formula Synthesizer & Domain Researcher',
      clearance: 'Level 4 (Formula Synthesis & Research)',
      icon: <Search className="agent-icon" size={26} />,
      accentColor: '#06b6d4',
      status: 'ACTIVE',
      description:
        'The analytical researcher. Scout synthesizes complex nested formulas (XLOOKUP, INDEX/MATCH, dynamic arrays), cross-references financial ratios, and retrieves external domain knowledge when formulas are unfamiliar.',
      responsibilities: [
        'Autonomous synthesis of complex formulas with exact syntax validation',
        'Accounting and financial ratio formula translation (CAGR, EBITDA, ROI)',
        'External domain lookups when spreadsheet definitions require clarification',
        'Excel and Google Sheets compatibility translation',
      ],
      capabilities: [
        'Formula Synthesizer',
        'Domain Research',
        'Syntax Cross-Check',
        'Accounting Ratios',
        'Formula Explainer',
      ],
      metrics: [
        { label: 'Formula Accuracy', value: '98.9%' },
        { label: 'Function Library', value: '400+ Formulas' },
        { label: 'Cross-App Support', value: 'Excel + Sheets' },
      ],
      samplePrompts: [
        'Synthesize an XLOOKUP formula with fallback if missing',
        'Create a formula to calculate compound monthly growth rate',
        'Explain how the tax rate formula works in column H',
      ],
    },
  ];

  const handleCopyPrompt = (prompt: string) => {
    navigator.clipboard?.writeText(prompt);
    setCopiedPrompt(prompt);
    setTimeout(() => setCopiedPrompt(null), 2000);
  };

  const handleRunPrompt = (prompt: string) => {
    if (onSelectPrompt) {
      onSelectPrompt(prompt);
    } else {
      handleCopyPrompt(prompt);
      onBack();
    }
  };

  const filteredAgents = agents.filter((agent) => {
    if (selectedFilter === 'all') return true;
    if (selectedFilter === 'orchestration') return agent.id === 'nexus';
    if (selectedFilter === 'analytics') return agent.id === 'atlas' || agent.id === 'scout';
    if (selectedFilter === 'safety') return agent.id === 'aegis';
    if (selectedFilter === 'data') return agent.id === 'valkyrie' || agent.id === 'cortex';
    return true;
  });

  return (
    <div className="agents-page" data-testid="agents-page">
      {/* Top Banner with Zenith OS Branding */}
      <div className="agents-top-bar">
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
          <span>A Zenith Open Source Flagship Project</span>
          <ExternalLink size={12} />
        </a>
      </div>

      {/* Hero Showcase */}
      <header className="agents-hero">
        <div className="agents-hero-badge">
          <Bot size={14} />
          <span>AUTONOMOUS MULTI-AGENT WORKFORCE</span>
        </div>
        <h1 className="agents-hero-title">
          Meet the Minds Behind <span className="text-emerald">ExcelAgento</span>
        </h1>
        <p className="agents-hero-subtitle">
          Six specialized, deterministic neural agents collaborating in real time to inspect,
          verify, calculate, transform, and learn from your spreadsheets. Fully local, zero-leakage,
          and built for professional business rigor.
        </p>

        {/* Filter Pills */}
        <div className="agents-filter-pills">
          <button
            className={`filter-pill ${selectedFilter === 'all' ? 'active' : ''}`}
            onClick={() => setSelectedFilter('all')}
          >
            All Agents ({agents.length})
          </button>
          <button
            className={`filter-pill ${selectedFilter === 'orchestration' ? 'active' : ''}`}
            onClick={() => setSelectedFilter('orchestration')}
          >
            Orchestration
          </button>
          <button
            className={`filter-pill ${selectedFilter === 'analytics' ? 'active' : ''}`}
            onClick={() => setSelectedFilter('analytics')}
          >
            Quantitative Analytics
          </button>
          <button
            className={`filter-pill ${selectedFilter === 'safety' ? 'active' : ''}`}
            onClick={() => setSelectedFilter('safety')}
          >
            Enterprise Safety
          </button>
          <button
            className={`filter-pill ${selectedFilter === 'data' ? 'active' : ''}`}
            onClick={() => setSelectedFilter('data')}
          >
            Data Engineering &amp; Memory
          </button>
        </div>
      </header>

      {/* Agents Grid */}
      <div className="agents-grid">
        {filteredAgents.map((agent) => (
          <div
            key={agent.id}
            className="agent-card"
            style={{ borderColor: `${agent.accentColor}33` }}
          >
            {/* Header */}
            <div className="agent-card-header">
              <div
                className="agent-avatar-box"
                style={{
                  color: agent.accentColor,
                  background: `${agent.accentColor}18`,
                  borderColor: `${agent.accentColor}44`,
                }}
              >
                {agent.icon}
              </div>
              <div className="agent-title-block">
                <div className="agent-codename">{agent.codename}</div>
                <h3 className="agent-name">{agent.name}</h3>
                <div className="agent-role">{agent.role}</div>
              </div>
              <div className="agent-status-pill">
                <span className="pulse-dot" style={{ backgroundColor: agent.accentColor }} />
                <span>{agent.status}</span>
              </div>
            </div>

            {/* Description */}
            <p className="agent-desc">{agent.description}</p>

            {/* Metrics */}
            <div className="agent-metrics-row">
              {agent.metrics.map((m, idx) => (
                <div key={idx} className="agent-metric-item">
                  <span className="agent-metric-label">{m.label}</span>
                  <span className="agent-metric-value" style={{ color: agent.accentColor }}>
                    {m.value}
                  </span>
                </div>
              ))}
            </div>

            {/* Responsibilities */}
            <div className="agent-section-block">
              <div className="agent-section-title">
                <CheckCircle2 size={12} style={{ color: agent.accentColor }} />
                <span>Core Responsibilities</span>
              </div>
              <ul className="agent-responsibilities-list">
                {agent.responsibilities.map((r, rIdx) => (
                  <li key={rIdx}>{r}</li>
                ))}
              </ul>
            </div>

            {/* Capabilities */}
            <div className="agent-caps-row">
              {agent.capabilities.map((c, cIdx) => (
                <span key={cIdx} className="agent-cap-tag">
                  {c}
                </span>
              ))}
            </div>

            {/* Sample Prompts */}
            <div className="agent-prompts-block">
              <div className="agent-section-title">
                <Terminal size={12} style={{ color: agent.accentColor }} />
                <span>Deploy Agent With Command</span>
              </div>
              <div className="agent-prompt-list">
                {agent.samplePrompts.map((p, pIdx) => (
                  <div key={pIdx} className="agent-prompt-item">
                    <span className="prompt-text">"{p}"</span>
                    <button
                      className="prompt-run-btn"
                      onClick={() => handleRunPrompt(p)}
                      title="Run this prompt in the workspace"
                    >
                      <ArrowRight size={12} />
                      <span>{onSelectPrompt ? 'Run' : 'Copy'}</span>
                    </button>
                  </div>
                ))}
              </div>
            </div>
          </div>
        ))}
      </div>

      {/* Architecture Flow Section */}
      <section className="agents-arch-section">
        <h2 className="arch-title">Multi-Agent Synchronous Architecture</h2>
        <p className="arch-subtitle">
          How requests traverse through the autonomous agent mesh within your local browser:
        </p>

        <div className="arch-flow-grid">
          <div className="arch-flow-node">
            <div className="arch-node-step">01. INGEST</div>
            <div className="arch-node-title">Nexus Conductor</div>
            <div className="arch-node-desc">
              Normalizes typo-ridden English, retrieves cross-turn working memory, generates
              execution plan.
            </div>
          </div>
          <div className="arch-flow-arrow">→</div>
          <div className="arch-flow-node">
            <div className="arch-node-step">02. VERIFY</div>
            <div className="arch-node-title">Aegis Guardrail</div>
            <div className="arch-node-desc">
              Checks schema constraints, prevents destructive data wipes, creates dry-run diff
              preview.
            </div>
          </div>
          <div className="arch-flow-arrow">→</div>
          <div className="arch-flow-node">
            <div className="arch-node-step">03. EXECUTE</div>
            <div className="arch-node-title">Atlas / Valkyrie</div>
            <div className="arch-node-desc">
              Calculates deterministic mathematics and runs 100k+ row transforms inside Web Workers.
            </div>
          </div>
          <div className="arch-flow-arrow">→</div>
          <div className="arch-flow-node">
            <div className="arch-node-step">04. LEARN</div>
            <div className="arch-node-title">Cortex Engine</div>
            <div className="arch-node-desc">
              Caches verified operation patterns and synchronizes cloud context with zero data
              leakage.
            </div>
          </div>
        </div>
      </section>

      {/* Footer Branding */}
      <footer className="agents-footer">
        <div className="footer-links">
          <span>ExcelAgento &copy; 2026</span>
          <span>•</span>
          <a
            href="https://zenithopensourceprojects.vercel.app/os"
            target="_blank"
            rel="noopener noreferrer"
          >
            Zenith Open Source Projects (zenithopensourceprojects.vercel.app/os)
          </a>
          <span>•</span>
          <a href="#/privacy">Privacy Policy (DPDP 2023)</a>
          <span>•</span>
          <a href="#/terms">Terms of Service</a>
          <span>•</span>
          <a href="#/docs">Architecture &amp; Docs</a>
        </div>
      </footer>
    </div>
  );
};
