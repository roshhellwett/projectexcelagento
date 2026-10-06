import React from 'react';
import { motion, useReducedMotion } from 'framer-motion';
import {
  ArrowRight,
  Check,
  ChevronRight,
  CircleCheck,
  Clock3,
  FileSpreadsheet,
  LockKeyhole,
  Orbit,
  ScanSearch,
  ShieldCheck,
  SlidersHorizontal,
  Sparkles,
  Undo2,
} from 'lucide-react';
import { useAuth } from '../lib/auth-context.js';

interface LandingPageProps {
  onLaunchWorkspace: () => void;
  onOpenAuth: (mode?: 'signin' | 'signup') => void;
  onOpenAgents: () => void;
  onOpenDocs: () => void;
  onOpenPrivacy: () => void;
}

const capabilities = [
  {
    icon: ScanSearch,
    label: 'Inspect before acting',
    title: 'The request starts with the workbook.',
    description:
      'The agent profiles the active sheet, reads the relevant structure, and grounds its response in what is actually present.',
  },
  {
    icon: SlidersHorizontal,
    label: 'Schema-validated operations',
    title: 'Proposals become real engine operations.',
    description:
      'Supported transformations are validated against the workbook before they are shown as ready to review.',
  },
  {
    icon: ShieldCheck,
    label: 'Human review stays in the loop',
    title: 'You decide what reaches the sheet.',
    description:
      'Previews show the affected range. Destructive changes require explicit confirmation, and every commit is reversible.',
  },
  {
    icon: LockKeyhole,
    label: 'Browser-first by design',
    title: 'Local work stays local by default.',
    description:
      'Parsing, deterministic analysis, history, checkpoints, and verified memory run in the browser. Model access is optional BYOK.',
  },
];

const workflowStages = [
  [
    '01',
    'Understand',
    'Interpret the instruction and ask for clarification when the goal is ambiguous.',
  ],
  ['02', 'Inspect', 'Read the workbook shape and the bounded context needed for the request.'],
  [
    '03',
    'Plan',
    'Turn the goal into one operation or a staged execution plan using the engine catalog.',
  ],
  [
    '04',
    'Verify',
    'Validate every step, calculate a preview, and block unsafe or unsupported changes.',
  ],
  [
    '05',
    'Apply',
    'Commit once, record the result, and keep undo/redo available at the same boundary.',
  ],
];

const specialistStages = [
  ['Conductor', 'Breaks a complex request into bounded work.'],
  ['Analyst', 'Grounds the request with read-only workbook tools.'],
  ['Planner', 'Drafts a strict plan from cataloged operations.'],
  ['Critic', 'Reviews the plan and can request one bounded revision.'],
  ['Verifier', 'Validates schema, arguments, previews, and safety gates.'],
];

const operationTags = [
  'Clean duplicates',
  'Normalize text',
  'Format dates',
  'Sort and filter',
  'Fill blanks',
  'Join sheets',
  'Group and summarize',
  'Find outliers',
  'Compare workbook states',
];

function Brand() {
  return (
    <span className="product-brand">
      <span className="product-brand-mark">
        <img src="/excel-agent-logo.svg" alt="" />
      </span>
      <span>
        <strong>
          Excel<span>Agento</span>
        </strong>
        <small>THE AGENT WORKSPACE</small>
      </span>
    </span>
  );
}

export const LandingPage: React.FC<LandingPageProps> = ({
  onLaunchWorkspace,
  onOpenAuth,
  onOpenAgents,
  onOpenDocs,
  onOpenPrivacy,
}) => {
  const { user, profile, signOut } = useAuth();
  const reducedMotion = useReducedMotion();

  const goToSection = (id: string) => {
    document.getElementById(id)?.scrollIntoView({ behavior: reducedMotion ? 'auto' : 'smooth' });
  };

  const openWorkspace = () => {
    if (user) {
      onLaunchWorkspace();
    } else {
      onOpenAuth('signup');
    }
  };

  return (
    <div className="product-page">
      <div className="product-page-grid" aria-hidden="true" />
      <header className="product-nav">
        <div className="product-nav-inner">
          <button
            type="button"
            className="product-brand-button"
            onClick={() => goToSection('product-top')}
          >
            <Brand />
          </button>

          <nav className="product-nav-links" aria-label="Product navigation">
            <button type="button" onClick={() => goToSection('capabilities')}>
              Capabilities
            </button>
            <button type="button" onClick={() => goToSection('workflow')}>
              How it works
            </button>
            <button type="button" onClick={onOpenAgents}>
              Agents
            </button>
            <button type="button" onClick={onOpenDocs}>
              Docs
            </button>
          </nav>

          <div className="product-nav-actions">
            {user ? (
              <>
                <span className="product-account-label">
                  <span className="product-account-dot" />
                  {profile?.full_name || user.email?.split('@')[0] || 'Account'}
                </span>
                <button
                  type="button"
                  className="product-button product-button-quiet"
                  onClick={() => void signOut()}
                >
                  Sign out
                </button>
                <button
                  type="button"
                  className="product-button product-button-primary"
                  onClick={onLaunchWorkspace}
                >
                  Open workspace <ArrowRight size={15} />
                </button>
              </>
            ) : (
              <>
                <button
                  type="button"
                  className="product-button product-button-quiet"
                  onClick={() => onOpenAuth('signin')}
                >
                  Sign in
                </button>
                <button
                  type="button"
                  className="product-button product-button-primary"
                  onClick={() => onOpenAuth('signup')}
                >
                  Create account <ArrowRight size={15} />
                </button>
              </>
            )}
          </div>
        </div>
      </header>

      <main>
        <section id="product-top" className="product-hero product-container">
          <motion.div
            className="product-hero-copy"
            initial={reducedMotion ? false : { opacity: 0, y: 18 }}
            animate={reducedMotion ? undefined : { opacity: 1, y: 0 }}
            transition={{ duration: 0.5, ease: [0.2, 0.7, 0.3, 1] }}
          >
            <div className="product-eyebrow">
              <span /> OPEN-SOURCE SPREADSHEET WORKSPACE
            </div>
            <h1>Tell your workbook what needs to change.</h1>
            <p className="product-hero-lede">
              ExcelAgento turns plain-English requests into grounded, reviewable spreadsheet work.
              Load a workbook, describe the outcome, inspect the proposed operation, and keep the
              final decision in your hands.
            </p>
            <div className="product-hero-actions">
              <button
                type="button"
                className="product-button product-button-primary product-button-large"
                onClick={openWorkspace}
              >
                {user ? 'Open workspace' : 'Create account to enter'} <ArrowRight size={17} />
              </button>
              <button
                type="button"
                className="product-button product-button-secondary product-button-large"
                onClick={() => goToSection('workflow')}
              >
                See the workflow <ChevronRight size={16} />
              </button>
            </div>
            <p className="product-access-note">
              <LockKeyhole size={13} /> An account is required before entering the workspace. One
              open-source workspace experience.
            </p>
          </motion.div>

          <motion.div
            className="product-console"
            initial={reducedMotion ? false : { opacity: 0, y: 24, scale: 0.98 }}
            animate={reducedMotion ? undefined : { opacity: 1, y: 0, scale: 1 }}
            transition={{ delay: 0.08, duration: 0.55, ease: [0.2, 0.7, 0.3, 1] }}
            aria-label="Illustration of the ExcelAgento workspace"
          >
            <div className="product-console-topbar">
              <div className="product-console-dots">
                <span />
                <span />
                <span />
              </div>
              <span>
                <FileSpreadsheet size={13} /> workspace interface
              </span>
              <small>WORKSPACE PREVIEW</small>
            </div>
            <div className="product-console-body">
              <div className="product-console-toolbar">
                <span className="product-console-file">
                  <FileSpreadsheet size={14} /> your workbook
                </span>
                <span className="product-console-chip">Active sheet</span>
                <span className="product-console-chip">Local engine</span>
              </div>
              <div className="product-console-content">
                <div className="product-mini-grid">
                  <div className="product-mini-grid-head">
                    <span>#</span>
                    <span>Workspace surface</span>
                    <span>State</span>
                    <span> </span>
                  </div>
                  <div>
                    <span>01</span>
                    <span>Workbook loaded</span>
                    <span className="is-muted">ready</span>
                    <span />
                  </div>
                  <div>
                    <span>02</span>
                    <span>Plain-English request</span>
                    <span className="is-active">review</span>
                    <span />
                  </div>
                  <div>
                    <span>03</span>
                    <span>Bounded preview</span>
                    <span className="is-muted">ready</span>
                    <span />
                  </div>
                  <div>
                    <span>04</span>
                    <span>Apply and history</span>
                    <span className="is-muted">controlled</span>
                    <span />
                  </div>
                </div>
                <div className="product-agent-card">
                  <div className="product-agent-card-head">
                    <span className="product-status-pulse" /> AGENT PROPOSAL <span>PREVIEW</span>
                  </div>
                  <p>
                    “Normalize the dates in this sheet and keep the original values available for
                    review.”
                  </p>
                  <div className="product-agent-step">
                    <CircleCheck size={14} /> Inspect active sheet
                  </div>
                  <div className="product-agent-step">
                    <CircleCheck size={14} /> Prepare a bounded operation
                  </div>
                  <button type="button" className="product-preview-button" onClick={openWorkspace}>
                    Review in workspace <ArrowRight size={13} />
                  </button>
                </div>
              </div>
            </div>
          </motion.div>
        </section>

        <section className="product-proof-strip product-container" aria-label="Product principles">
          <div>
            <Orbit size={16} />
            <span>Browser-first execution</span>
          </div>
          <div>
            <ShieldCheck size={16} />
            <span>Review before apply</span>
          </div>
          <div>
            <Undo2 size={16} />
            <span>Reversible history</span>
          </div>
          <div>
            <Sparkles size={16} />
            <span>Optional BYOK reasoning</span>
          </div>
        </section>

        <section id="capabilities" className="product-section product-container">
          <div className="product-section-heading">
            <span className="product-section-kicker">WHAT IS ACTUALLY HERE</span>
            <h2>A careful layer between your instruction and your workbook.</h2>
            <p>
              The interface reflects the implementation: inspectable state, explicit proposals, and
              an engine that can say no when a request is unsupported or unsafe.
            </p>
          </div>
          <div className="product-capability-grid">
            {capabilities.map(({ icon: Icon, label, title, description }, index) => (
              <motion.article
                key={label}
                className="product-capability-card"
                initial={reducedMotion ? false : { opacity: 0, y: 14 }}
                whileInView={reducedMotion ? undefined : { opacity: 1, y: 0 }}
                viewport={{ once: true, amount: 0.25 }}
                transition={{ delay: index * 0.04, duration: 0.35 }}
              >
                <div className="product-card-icon">
                  <Icon size={18} />
                </div>
                <span className="product-card-label">{label}</span>
                <h3>{title}</h3>
                <p>{description}</p>
              </motion.article>
            ))}
          </div>
        </section>

        <section id="workflow" className="product-section product-section-tinted">
          <div className="product-container">
            <div className="product-section-heading product-section-heading-left">
              <span className="product-section-kicker">A REQUEST HAS A TRACE</span>
              <h2>From plain English to a verified commit.</h2>
              <p>
                Simple requests stay on the fast local path. Complex requests can move through the
                bounded multi-agent pipeline before the workspace offers an apply action.
              </p>
            </div>
            <div className="product-workflow">
              {workflowStages.map(([number, title, description]) => (
                <div key={number} className="product-workflow-step">
                  <span className="product-workflow-number">{number}</span>
                  <div>
                    <h3>{title}</h3>
                    <p>{description}</p>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </section>

        <section className="product-section product-container product-swarm-section">
          <div className="product-swarm-intro">
            <span className="product-section-kicker">WHEN THE REQUEST IS COMPLEX</span>
            <h2>Specialists collaborate behind one accountable interface.</h2>
            <p>
              The agents are roles in a controlled pipeline, not a collection of decorative
              personas. Each stage has a bounded job and the verifier is the only path to an
              apply-ready plan.
            </p>
            <button type="button" className="product-inline-link" onClick={onOpenAgents}>
              Explore the agent architecture <ArrowRight size={15} />
            </button>
          </div>
          <div className="product-specialist-list">
            {specialistStages.map(([name, description], index) => (
              <div key={name} className="product-specialist-row">
                <span>{String(index + 1).padStart(2, '0')}</span>
                <strong>{name}</strong>
                <p>{description}</p>
                <Check size={15} />
              </div>
            ))}
          </div>
        </section>

        <section className="product-section product-section-tinted product-operations-section">
          <div className="product-container product-operations-layout">
            <div>
              <span className="product-section-kicker">WORKBOOK WORK, NOT PROMISES</span>
              <h2>Useful for the operations people repeat every week.</h2>
              <p>
                Ask for supported transformations, calculations, or analysis. The operation catalog,
                validation layer, and usage ledger keep the result legible.
              </p>
            </div>
            <div className="product-operation-tags">
              {operationTags.map((tag) => (
                <span key={tag}>{tag}</span>
              ))}
            </div>
          </div>
        </section>

        <section className="product-section product-container product-data-section">
          <div className="product-data-card">
            <div className="product-data-card-icon">
              <LockKeyhole size={20} />
            </div>
            <div>
              <span className="product-section-kicker">DATA FLOW, PLAINLY STATED</span>
              <h2>Your workbook is local by default. Model access is your choice.</h2>
              <p>
                Without an API key, the deterministic planner, formula-aware analysis, workbook
                parsing, checkpoints, and verified memory run in this browser. If you connect Groq,
                OpenRouter, Gemini, OpenAI, or another compatible provider, your browser sends the
                prompt and the bounded context required for that request directly to the provider.
              </p>
              <button type="button" className="product-inline-link" onClick={onOpenPrivacy}>
                Read the data-flow details <ArrowRight size={15} />
              </button>
            </div>
          </div>
        </section>

        <section className="product-final-cta product-container">
          <div>
            <span className="product-section-kicker">READY WHEN YOU ARE</span>
            <h2>Bring a real workbook. Leave with a traceable result.</h2>
            <p>
              Create an account to enter the workspace. The project is open source and has one
              workspace experience.
            </p>
          </div>
          <button
            type="button"
            className="product-button product-button-primary product-button-large"
            onClick={openWorkspace}
          >
            {user ? 'Open workspace' : 'Create account'} <ArrowRight size={17} />
          </button>
        </section>
      </main>

      <footer className="product-footer">
        <div className="product-container product-footer-inner">
          <div>
            <Brand />
            <p>Plain-English spreadsheet work with visible boundaries.</p>
          </div>
          <div className="product-footer-links">
            <button type="button" onClick={onOpenDocs}>
              Documentation
            </button>
            <button type="button" onClick={onOpenAgents}>
              Agent architecture
            </button>
            <button type="button" onClick={onOpenPrivacy}>
              Privacy and data flow
            </button>
          </div>
          <div className="product-footer-status">
            <Clock3 size={14} /> Open source · account required for workspace access
          </div>
        </div>
      </footer>
    </div>
  );
};
