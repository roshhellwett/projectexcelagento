import React, { useState } from 'react';
import {
  Brain,
  Shield,
  Zap,
  Cpu,
  Sparkles,
  Lock,
  ArrowRight,
  BarChart3,
  Code2,
  FileSpreadsheet,
  TrendingUp,
  UserCheck,
  Check,
} from 'lucide-react';
import { useAuth } from '../lib/auth-context.js';

interface LandingPageProps {
  onLaunchWorkspace: () => void;
  onOpenAuth: (mode?: 'signin' | 'signup') => void;
  onOpenAgents: () => void;
  onOpenDocs: () => void;
  onOpenPrivacy: () => void;
}

export const LandingPage: React.FC<LandingPageProps> = ({
  onLaunchWorkspace,
  onOpenAuth,
  onOpenAgents,
  onOpenDocs,
  onOpenPrivacy,
}) => {
  const { user, profile, signOut } = useAuth();
  const [billingCycle, setBillingCycle] = useState<'monthly' | 'annual'>('monthly');

  const agents = [
    {
      id: 'nexus',
      name: 'Nexus Conductor',
      codename: 'NEXUS-01',
      role: 'Intent Classification & Phased Planning',
      icon: <Brain size={20} className="text-emerald" />,
      color: 'emerald',
      description: 'Parses messy natural language, classifies user intent, and plans multi-step execution graphs.',
    },
    {
      id: 'atlas',
      name: 'Atlas Scientist',
      codename: 'ATLAS-02',
      role: 'Deterministic Financial Math & Analytics',
      icon: <BarChart3 size={20} className="text-cyan" />,
      color: 'cyan',
      description: 'Executes 64-bit IEEE-754 calculations, CAGR, variance, tax brackets, and statistical distributions without LLM math hallucinations.',
    },
    {
      id: 'aegis',
      name: 'Aegis Guardrail',
      codename: 'AEGIS-03',
      role: 'Safety Verification & Invariant Gatekeeping',
      icon: <Shield size={20} className="text-amber" />,
      color: 'amber',
      description: 'Validates schemas before and after changes. Demands explicit confirmation before destructive row or column modifications.',
    },
    {
      id: 'cortex',
      name: 'Cortex Neural',
      codename: 'CORTEX-04',
      role: '100% In-Browser Local Memory',
      icon: <Cpu size={20} className="text-purple" />,
      color: 'purple',
      description: 'Maintains verified formula associations and multi-turn working scratchpads locally at 0ms latency with zero network roundtrips.',
    },
    {
      id: 'valkyrie',
      name: 'Valkyrie Engine',
      codename: 'VALKYRIE-05',
      role: 'Non-Blocking Big Data Web Worker',
      icon: <Zap size={20} className="text-blue" />,
      color: 'blue',
      description: 'Processes 150,000+ row datasets in dedicated background threads without causing browser frame drops or UI stutter.',
    },
    {
      id: 'scout',
      name: 'Scout Knowledge',
      codename: 'SCOUT-06',
      role: 'Autonomous Formula Synthesis',
      icon: <Code2 size={20} className="text-rose" />,
      color: 'rose',
      description: 'Synthesizes and repairs complex modern formulas (XLOOKUP, INDEX/MATCH, Dynamic Arrays) and checks dependency DAGs.',
    },
  ];

  return (
    <div className="landing-container">
      {/* Top Navigation */}
      <header className="landing-nav">
        <div className="landing-nav-inner">
          <div className="landing-brand">
            <span className="landing-logo-mark">Σ</span>
            <span className="landing-logo-name">ExcelAgento</span>
            <span className="landing-badge">Enterprise v0.2</span>
          </div>

          <nav className="landing-nav-links">
            <a href="#features" className="landing-nav-link">Features</a>
            <a href="#swarm" className="landing-nav-link">Autonomous Swarm</a>
            <a href="#pricing" className="landing-nav-link">Pricing</a>
            <button type="button" onClick={onOpenDocs} className="landing-nav-link btn-link">
              Docs
            </button>
            <button type="button" onClick={onOpenAgents} className="landing-nav-link btn-link">
              Agents
            </button>
          </nav>

          <div className="landing-nav-actions">
            {user ? (
              <div className="landing-user-menu">
                <span className="landing-user-pill">
                  <UserCheck size={14} className="text-emerald" />
                  <span className="landing-user-name">
                    {profile?.full_name || user.email?.split('@')[0]}
                  </span>
                  <span className="landing-tier-pill">
                    {profile?.tier?.toUpperCase() || 'FREE'}
                  </span>
                </span>
                <button
                  type="button"
                  className="btn btn-secondary btn-sm"
                  onClick={onLaunchWorkspace}
                >
                  Workspace
                </button>
                <button
                  type="button"
                  className="btn btn-ghost btn-sm"
                  onClick={() => signOut()}
                >
                  Sign Out
                </button>
              </div>
            ) : (
              <>
                <button
                  type="button"
                  className="btn btn-ghost btn-sm"
                  onClick={() => onOpenAuth('signin')}
                >
                  Sign In
                </button>
                <button
                  type="button"
                  className="btn btn-primary btn-sm"
                  onClick={onLaunchWorkspace}
                >
                  <span>Launch Workspace</span>
                  <ArrowRight size={14} />
                </button>
              </>
            )}
          </div>
        </div>
      </header>

      {/* Hero Section */}
      <section className="landing-hero">
        <div className="landing-hero-backdrop" />
        <div className="landing-hero-content">
          <div className="landing-eyebrow">
            <Sparkles size={14} className="text-emerald" />
            <span>Autonomous In-Browser Spreadsheet Intelligence</span>
          </div>

          <h1 className="landing-headline">
            Autonomous AI Agents for Spreadsheets.<br />
            <span className="landing-headline-gradient">100% In-Browser. 0ms Latency.</span>
          </h1>

          <p className="landing-subtext">
            Deterministic 64-bit mathematical precision meets 6 collaborative autonomous agents.
            Perform automated data cleaning, complex formula synthesis, and big data transformations
            with zero external database transmission.
          </p>

          <div className="landing-hero-actions">
            <button
              type="button"
              className="btn btn-primary btn-lg landing-cta-btn"
              onClick={onLaunchWorkspace}
            >
              <span>Launch Free Workspace</span>
              <ArrowRight size={18} />
            </button>
            <button
              type="button"
              className="btn btn-secondary btn-lg"
              onClick={onOpenAgents}
            >
              <Brain size={18} className="text-emerald" />
              <span>Explore The Swarm</span>
            </button>
          </div>

          {/* Metric Bar */}
          <div className="landing-metrics-bar">
            <div className="landing-metric-item">
              <span className="landing-metric-val">0ms</span>
              <span className="landing-metric-label">Memory Latency</span>
            </div>
            <div className="landing-metric-divider" />
            <div className="landing-metric-item">
              <span className="landing-metric-val">100%</span>
              <span className="landing-metric-label">Client-Side Privacy</span>
            </div>
            <div className="landing-metric-divider" />
            <div className="landing-metric-item">
              <span className="landing-metric-val">150,000+</span>
              <span className="landing-metric-label">Row Capacity</span>
            </div>
            <div className="landing-metric-divider" />
            <div className="landing-metric-item">
              <span className="landing-metric-val">6</span>
              <span className="landing-metric-label">Specialized Agents</span>
            </div>
          </div>
        </div>

        {/* Hero Interactive Mockup Banner */}
        <div className="landing-hero-banner-wrapper">
          <div className="landing-hero-banner">
            <div className="landing-banner-header">
              <div className="landing-window-dots">
                <span className="dot red" />
                <span className="dot yellow" />
                <span className="dot green" />
              </div>
              <div className="landing-banner-title">
                <FileSpreadsheet size={14} className="text-emerald" />
                <span>Q3_Financial_Performance_Model.xlsx</span>
                <span className="landing-sheet-badge">Active</span>
              </div>
              <div className="landing-banner-meta">
                <span className="meta-pill">12,450 Rows</span>
                <span className="meta-pill">64-bit IEEE</span>
              </div>
            </div>

            {/* Simulated Workspace Grid & Swarm HUD */}
            <div className="landing-banner-body">
              {/* Swarm HUD Bar */}
              <div className="landing-hud-strip">
                <div className="hud-badge is-active">
                  <span className="hud-pulse" />
                  <Brain size={12} className="text-emerald" />
                  <span>NEXUS: Orchestrating</span>
                </div>
                <div className="hud-badge is-active">
                  <span className="hud-pulse cyan" />
                  <BarChart3 size={12} className="text-cyan" />
                  <span>ATLAS: CAGR 14.8%</span>
                </div>
                <div className="hud-badge">
                  <Shield size={12} className="text-amber" />
                  <span>AEGIS: Invariant Guard</span>
                </div>
                <div className="hud-badge">
                  <Cpu size={12} className="text-purple" />
                  <span>CORTEX: 0ms Local Memory</span>
                </div>
              </div>

              {/* Data Table Preview */}
              <div className="landing-grid-preview">
                <div className="preview-row header">
                  <span className="cell index">#</span>
                  <span className="cell">Fiscal Period</span>
                  <span className="cell">Region</span>
                  <span className="cell">Revenue (USD)</span>
                  <span className="cell">Operating Margin</span>
                  <span className="cell">CAGR Formula</span>
                </div>
                <div className="preview-row">
                  <span className="cell index">1</span>
                  <span className="cell">FY 2023 - Q1</span>
                  <span className="cell">North America</span>
                  <span className="cell font-mono">$1,420,500</span>
                  <span className="cell font-mono text-emerald">24.2%</span>
                  <span className="cell formula">=RATE(4, 0, -C2, C5)</span>
                </div>
                <div className="preview-row">
                  <span className="cell index">2</span>
                  <span className="cell">FY 2023 - Q2</span>
                  <span className="cell">North America</span>
                  <span className="cell font-mono">$1,580,200</span>
                  <span className="cell font-mono text-emerald">25.6%</span>
                  <span className="cell formula">=XLOOKUP(A3, Targets!A:A, Targets!B:B)</span>
                </div>
                <div className="preview-row highlight">
                  <span className="cell index">3</span>
                  <span className="cell">FY 2023 - Q3</span>
                  <span className="cell">Europe & APAC</span>
                  <span className="cell font-mono">$2,190,000</span>
                  <span className="cell font-mono text-emerald">28.4%</span>
                  <span className="cell formula text-cyan font-bold">=SUMIFS(E:E, B:B, "APAC")</span>
                </div>
              </div>

              {/* Reasoning Card Callout */}
              <div className="landing-agent-callout">
                <div className="agent-callout-header">
                  <Sparkles size={14} className="text-emerald" />
                  <span className="font-semibold">Atlas Scientist Verification</span>
                  <span className="agent-callout-tag">Deterministic 0ms</span>
                </div>
                <p className="agent-callout-text">
                  Evaluated 12,450 rows. Variance calculation confirmed within 0.0001% tolerance.
                  Proposed formula <code>=SUMIFS(E:E, B:B, &quot;APAC&quot;)</code> passed Aegis destructive guardrail inspection.
                </p>
              </div>
            </div>
          </div>
        </div>
      </section>

      {/* Feature Pillars */}
      <section id="features" className="landing-section">
        <div className="landing-section-header">
          <span className="landing-section-pill">Precision Architecture</span>
          <h2 className="landing-section-title">Engineered For Spreadsheet Experts</h2>
          <p className="landing-section-sub">
            Say goodbye to fragile cloud AI extensions that hallucinate spreadsheet formulas or leak confidential company financial records.
          </p>
        </div>

        <div className="landing-features-grid">
          <div className="landing-feature-card">
            <div className="feature-icon-box bg-emerald-dim">
              <Lock size={22} className="text-emerald" />
            </div>
            <h3>100% In-Browser Privacy</h3>
            <p>
              Every row, cell, and formula is computed in local browser memory and Web Workers.
              Zero worksheet data is transmitted to external cloud databases.
            </p>
          </div>

          <div className="landing-feature-card">
            <div className="feature-icon-box bg-cyan-dim">
              <TrendingUp size={22} className="text-cyan" />
            </div>
            <h3>Deterministic 64-Bit Math</h3>
            <p>
              Calculations are executed with strict IEEE-754 precision. No hallucinated averages or fabricated financial numbers.
            </p>
          </div>

          <div className="landing-feature-card">
            <div className="feature-icon-box bg-purple-dim">
              <Zap size={22} className="text-purple" />
            </div>
            <h3>0ms Local Memory</h3>
            <p>
              Self-learning verified action memory operates at 0ms latency in browser memory. No network hops during real-time reasoning turns.
            </p>
          </div>

          <div className="landing-feature-card">
            <div className="feature-icon-box bg-amber-dim">
              <Shield size={22} className="text-amber" />
            </div>
            <h3>Aegis Invariant Gatekeeper</h3>
            <p>
              Automated schema validation prevents data corruption. Structural changes demand explicit confirmation before applying.
            </p>
          </div>
        </div>
      </section>

      {/* Swarm Intelligence Section */}
      <section id="swarm" className="landing-section bg-alt">
        <div className="landing-section-header">
          <span className="landing-section-pill">Multi-Agent Hive</span>
          <h2 className="landing-section-title">The 6 Autonomous Specialists</h2>
          <p className="landing-section-sub">
            Complex workflows require division of labor. Six specialized agents coordinate in real time to understand, calculate, verify, and execute.
          </p>
        </div>

        <div className="landing-agents-grid">
          {agents.map((agent) => (
            <div key={agent.id} className="landing-agent-card">
              <div className="agent-card-top">
                <div className={`agent-icon-avatar bg-${agent.color}-dim`}>
                  {agent.icon}
                </div>
                <div className="agent-meta">
                  <h4 className="agent-name">{agent.name}</h4>
                  <span className="agent-codename">{agent.codename}</span>
                </div>
              </div>
              <span className="agent-role-pill">{agent.role}</span>
              <p className="agent-desc">{agent.description}</p>
            </div>
          ))}
        </div>
      </section>

      {/* Pricing & Subscription Architecture Section */}
      <section id="pricing" className="landing-section">
        <div className="landing-section-header">
          <span className="landing-section-pill">Transparent Plans</span>
          <h2 className="landing-section-title">Built For Individuals &amp; Teams</h2>
          <p className="landing-section-sub">
            Start free with full local computing. Upgrade for high-volume reasoning credits, dedicated workers, and team collaboration.
          </p>

          <div className="landing-billing-toggle">
            <button
              type="button"
              className={`toggle-opt ${billingCycle === 'monthly' ? 'is-active' : ''}`}
              onClick={() => setBillingCycle('monthly')}
            >
              Monthly
            </button>
            <button
              type="button"
              className={`toggle-opt ${billingCycle === 'annual' ? 'is-active' : ''}`}
              onClick={() => setBillingCycle('annual')}
            >
              Annual <span className="discount-badge">Save 20%</span>
            </button>
          </div>
        </div>

        <div className="landing-pricing-grid">
          {/* Free Tier */}
          <div className="pricing-card">
            <div className="pricing-header">
              <h3 className="pricing-plan-name">Free Starter</h3>
              <p className="pricing-plan-desc">For individual analysts exploring autonomous spreadsheet intelligence.</p>
              <div className="pricing-price">
                <span className="price-val">$0</span>
                <span className="price-period">/ month</span>
              </div>
            </div>

            <ul className="pricing-features">
              <li><Check size={16} className="text-emerald" /> 50 AI Reasoning Credits / month</li>
              <li><Check size={16} className="text-emerald" /> 100% In-Browser Private Computing</li>
              <li><Check size={16} className="text-emerald" /> Access to all 6 Autonomous Agents</li>
              <li><Check size={16} className="text-emerald" /> Full 64-bit Spreadsheet Engine</li>
              <li><Check size={16} className="text-emerald" /> Browser-local Checkpoint Recovery</li>
              <li><Check size={16} className="text-emerald" /> Bring Your Own Key (BYOK) Support</li>
            </ul>

            <button
              type="button"
              className="btn btn-secondary btn-block"
              onClick={onLaunchWorkspace}
            >
              Start Free Now
            </button>
          </div>

          {/* Pro Tier */}
          <div className="pricing-card is-popular">
            <div className="popular-badge">Most Popular</div>
            <div className="pricing-header">
              <h3 className="pricing-plan-name">Pro Analyst</h3>
              <p className="pricing-plan-desc">For power users, financial analysts, and corporate modelers.</p>
              <div className="pricing-price">
                <span className="price-val">{billingCycle === 'annual' ? '$23' : '$29'}</span>
                <span className="price-period">/ month</span>
              </div>
            </div>

            <ul className="pricing-features">
              <li><Check size={16} className="text-emerald" /> <strong>1,000 AI Reasoning Credits / month</strong></li>
              <li><Check size={16} className="text-emerald" /> High-Concurrency Web Worker Acceleration</li>
              <li><Check size={16} className="text-emerald" /> 150,000+ Row Processing Capacity</li>
              <li><Check size={16} className="text-emerald" /> Unlimited Mission Undo / Redo History</li>
              <li><Check size={16} className="text-emerald" /> Priority Model Inference Streaming</li>
              <li><Check size={16} className="text-emerald" /> Advanced Financial &amp; Statistical Tool Suite</li>
              <li><Check size={16} className="text-emerald" /> Priority Support &amp; Feature Access</li>
            </ul>

            <button
              type="button"
              className="btn btn-primary btn-block"
              onClick={() => onOpenAuth('signup')}
            >
              Get Started with Pro
            </button>
          </div>

          {/* Enterprise Tier */}
          <div className="pricing-card">
            <div className="pricing-header">
              <h3 className="pricing-plan-name">Enterprise</h3>
              <p className="pricing-plan-desc">For organizations demanding custom security, SLAs, and unlimited scale.</p>
              <div className="pricing-price">
                <span className="price-val">{billingCycle === 'annual' ? '$79' : '$99'}</span>
                <span className="price-period">/ seat / mo</span>
              </div>
            </div>

            <ul className="pricing-features">
              <li><Check size={16} className="text-emerald" /> <strong>Unlimited AI Reasoning Credits</strong></li>
              <li><Check size={16} className="text-emerald" /> Custom Swarm Agent Training &amp; Domain Rules</li>
              <li><Check size={16} className="text-emerald" /> Dedicated SOC-2 &amp; DPDP Audit Exports</li>
              <li><Check size={16} className="text-emerald" /> Team Workspaces &amp; Shared Formula Cortex</li>
              <li><Check size={16} className="text-emerald" /> Custom LLM Base URL &amp; Self-Hosted Gateway</li>
              <li><Check size={16} className="text-emerald" /> 99.9% Uptime SLA &amp; Dedicated Engineer</li>
            </ul>

            <button
              type="button"
              className="btn btn-secondary btn-block"
              onClick={() => onOpenAuth('signup')}
            >
              Contact Enterprise
            </button>
          </div>
        </div>
      </section>

      {/* CTA Bottom Banner */}
      <section className="landing-cta-banner">
        <div className="cta-banner-card">
          <div className="cta-banner-text">
            <h2>Ready To Supercharge Your Workbooks?</h2>
            <p>Join thousands of financial modelers, data scientists, and analysts using ExcelAgento.</p>
          </div>
          <div className="cta-banner-buttons">
            <button
              type="button"
              className="btn btn-primary btn-lg"
              onClick={onLaunchWorkspace}
            >
              <span>Open Free Workspace</span>
              <ArrowRight size={18} />
            </button>
            <button
              type="button"
              className="btn btn-secondary btn-lg"
              onClick={() => onOpenAuth('signup')}
            >
              Create Account
            </button>
          </div>
        </div>
      </section>

      {/* Footer */}
      <footer className="landing-footer">
        <div className="landing-footer-inner">
          <div className="footer-brand-col">
            <div className="landing-brand">
              <span className="landing-logo-mark">Σ</span>
              <span className="landing-logo-name">ExcelAgento</span>
            </div>
            <p className="footer-desc">
              The autonomous in-browser spreadsheet copilot. 100% private, deterministic 64-bit math, and 0ms latency.
            </p>
            <span className="footer-statutory">Compliant with Indian DPDP Act 2023 Statutory Requirements</span>
          </div>

          <div className="footer-links-col">
            <h4>Product</h4>
            <button type="button" onClick={onLaunchWorkspace} className="footer-link">Workspace</button>
            <button type="button" onClick={onOpenAgents} className="footer-link">Autonomous Agents</button>
            <button type="button" onClick={onOpenDocs} className="footer-link">Developer Docs</button>
          </div>

          <div className="footer-links-col">
            <h4>Security &amp; Legal</h4>
            <button type="button" onClick={onOpenPrivacy} className="footer-link">Privacy Policy</button>
            <button type="button" onClick={onOpenDocs} className="footer-link">Security Disclosures</button>
            <a href="https://github.com" target="_blank" rel="noreferrer" className="footer-link">Open Source</a>
          </div>
        </div>

        <div className="footer-bottom">
          <span>&copy; {new Date().getFullYear()} ExcelAgento. All rights reserved.</span>
          <span>100% Client-Side Computing &bull; 0ms Local Memory</span>
        </div>
      </footer>
    </div>
  );
};
