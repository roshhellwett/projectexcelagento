import React, { useState } from 'react';
import {
  LifeBuoy,
  Send,
  CheckCircle2,
  AlertCircle,
  HelpCircle,
  Mail,
  ShieldCheck,
  Clock,
  Sparkles,
  ArrowRight,
} from 'lucide-react';
import { ResourcePageChrome, ResourcePageFooter } from './ResourcePageChrome.js';
import { submitSupportTicket } from '../lib/licensing.js';
import { useAuth } from '../lib/auth-context.js';

interface SupportPageProps {
  onBack: () => void;
  onNavigate?: (view: string) => void;
  onOpenAuth?: () => void;
  onOpenAdmin?: () => void;
}

const CATEGORIES = [
  { id: 'license', label: 'Access & Key Activation', icon: '🔑' },
  { id: 'bug', label: 'Bug or Calculation Error', icon: '🐛' },
  { id: 'feature', label: 'Feature Request / AI Capability', icon: '💡' },
  { id: 'billing', label: 'Commercial / Account Questions', icon: '💳' },
  { id: 'general', label: 'General Questions', icon: '💬' },
];

export const SupportPage: React.FC<SupportPageProps> = ({ onBack, onNavigate }) => {
  const { user } = useAuth();
  const [name, setName] = useState(user?.user_metadata?.full_name || '');
  const [email, setEmail] = useState(user?.email || '');
  const [category, setCategory] = useState('license');
  const [subject, setSubject] = useState('');
  const [message, setMessage] = useState('');
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState('');
  const [ticketId, setTicketId] = useState<string | null>(null);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!name.trim()) {
      setError('Please provide your name.');
      return;
    }
    if (!email.trim() || !email.includes('@')) {
      setError('Please provide a valid email address.');
      return;
    }
    if (!subject.trim()) {
      setError('Please enter an issue subject.');
      return;
    }
    if (!message.trim() || message.trim().length < 10) {
      setError('Please describe your issue in at least 10 characters.');
      return;
    }

    setSubmitting(true);
    setError('');

    try {
      const res = await submitSupportTicket({
        name,
        email,
        category,
        subject,
        message,
        userId: user?.id ?? null,
      });
      setTicketId(res.id);
    } catch (err) {
      setError(
        err instanceof Error ? err.message : 'Failed to submit support inquiry. Please try again.',
      );
    } finally {
      setSubmitting(false);
    }
  };

  const handleReset = () => {
    setTicketId(null);
    setSubject('');
    setMessage('');
    setError('');
  };

  return (
    <div className="resource-page support-page-surface">
      <ResourcePageChrome current="support" onBack={onBack} />

      <main className="resource-body support-body">
        {/* Hero Section */}
        <header className="resource-hero support-hero">
          <div className="support-hero-badge">
            <LifeBuoy size={16} />
            <span>EXCELAGENTO HELPDESK &amp; SUPPORT</span>
          </div>
          <h1>How can we help your spreadsheet workflows?</h1>
          <p>
            Every verified account starts with a 30-day evaluation. If access needs to continue, the
            operator can issue a server-validated activation key for a stated duration. Submit a
            request below for activation, account, calculation, or workflow help; this deployment
            has no in-app checkout or recurring card subscription.
          </p>
        </header>

        <div className="support-grid">
          {/* Main Form Column */}
          <section className="support-form-section">
            {ticketId ? (
              <div className="support-success-card">
                <div className="support-success-icon">
                  <CheckCircle2 size={42} className="text-emerald-500" />
                </div>
                <h2>Support Ticket Received!</h2>
                <p className="support-ticket-id">
                  Ticket Reference: <code>#TICK-{ticketId.slice(0, 8).toUpperCase()}</code>
                </p>
                <p className="support-success-desc">
                  We have logged your ticket into our operations queue. Our engineering and support
                  team will review your message and reply to <strong>{email}</strong>.
                </p>
                <div className="support-success-meta">
                  <div>
                    <Clock size={16} />
                    <span>
                      Reply destination: <strong>{email}</strong>
                    </span>
                  </div>
                  <div>
                    <ShieldCheck size={16} />
                    <span>Reference saved in database for tracking</span>
                  </div>
                </div>
                <div className="support-success-actions">
                  <button type="button" className="btn btn-secondary" onClick={handleReset}>
                    Submit another inquiry
                  </button>
                  <button type="button" className="btn btn-primary" onClick={onBack}>
                    Return to workspace
                  </button>
                </div>
              </div>
            ) : (
              <form className="support-card" onSubmit={handleSubmit}>
                <div className="support-card-header">
                  <div>
                    <span className="studio-eyebrow">DIRECT ASSISTANCE</span>
                    <h2>Submit a Support Request</h2>
                  </div>
                  <span className="support-priority-tag">Reply by email</span>
                </div>

                {error && (
                  <div className="support-alert danger" role="alert">
                    <AlertCircle size={18} />
                    <span>{error}</span>
                  </div>
                )}

                <div className="support-form-row">
                  <div className="support-field">
                    <label htmlFor="support-name">Your Full Name *</label>
                    <input
                      id="support-name"
                      type="text"
                      className="form-input"
                      placeholder="Jane Doe"
                      value={name}
                      onChange={(e) => setName(e.target.value)}
                      required
                    />
                  </div>
                  <div className="support-field">
                    <label htmlFor="support-email">Email Address *</label>
                    <input
                      id="support-email"
                      type="email"
                      className="form-input"
                      placeholder="jane@company.com"
                      value={email}
                      onChange={(e) => setEmail(e.target.value)}
                      required
                    />
                    <small className="field-hint">
                      We will reply directly to this email address.
                    </small>
                  </div>
                </div>

                <div className="support-field">
                  <label>Inquiry Category *</label>
                  <div className="support-category-grid">
                    {CATEGORIES.map((cat) => (
                      <button
                        key={cat.id}
                        type="button"
                        className={`support-cat-pill ${category === cat.id ? 'active' : ''}`}
                        onClick={() => setCategory(cat.id)}
                      >
                        <span className="cat-icon">{cat.icon}</span>
                        <span>{cat.label}</span>
                      </button>
                    ))}
                  </div>
                </div>

                <div className="support-field">
                  <label htmlFor="support-subject">Subject *</label>
                  <input
                    id="support-subject"
                    type="text"
                    className="form-input"
                    placeholder="e.g. License key expiration question or formula calculation bug"
                    value={subject}
                    onChange={(e) => setSubject(e.target.value)}
                    required
                  />
                </div>

                <div className="support-field">
                  <label htmlFor="support-message">Detailed Description *</label>
                  <textarea
                    id="support-message"
                    className="form-input support-textarea"
                    rows={6}
                    placeholder="Describe the issue or request in detail. If reporting a bug, please include steps to reproduce, sheet size, or error messages encountered."
                    value={message}
                    onChange={(e) => setMessage(e.target.value)}
                    required
                  />
                </div>

                <div className="support-submit-row">
                  <button
                    type="submit"
                    className="btn btn-primary support-submit-btn"
                    disabled={submitting}
                  >
                    {submitting ? (
                      <>
                        <span
                          className="spinner-border spinner-border-sm"
                          role="status"
                          aria-hidden="true"
                        />
                        Submitting ticket…
                      </>
                    ) : (
                      <>
                        <Send size={16} /> Submit Ticket
                      </>
                    )}
                  </button>
                  <span className="support-privacy-note">
                    Please do not include API keys or workbook contents. See Privacy &amp; data flow
                    for the implemented storage boundaries.
                  </span>
                </div>
              </form>
            )}
          </section>

          {/* Sidebar / FAQ Column */}
          <aside className="support-sidebar">
            <div className="support-card support-contact-card">
              <div className="support-card-header">
                <div>
                  <span className="studio-eyebrow">OFFICIAL CHANNELS</span>
                  <h3>Direct Contact</h3>
                </div>
                <Mail size={18} className="text-emerald-400" />
              </div>
              <p className="text-sm text-muted">
                Need an activation key, account-access help, or a commercial answer? Contact the
                operations inbox directly:
              </p>
              <div className="support-email-pill">
                <Mail size={14} />
                <a href="mailto:zenithprojects@icloud.com">zenithprojects@icloud.com</a>
              </div>
              <div className="support-sla-banner">
                <Sparkles size={15} />
                <span>Use this form or the operations email for account and access help.</span>
              </div>
            </div>

            <div className="support-card support-faq-card">
              <div className="support-card-header">
                <div>
                  <span className="studio-eyebrow">KNOWLEDGE BASE</span>
                  <h3>Quick Answers</h3>
                </div>
                <HelpCircle size={18} />
              </div>

              <div className="support-faq-list">
                <details className="faq-item" open>
                  <summary>How do I activate an activation key?</summary>
                  <p>
                    Go to <strong>Account &amp; Access</strong> in the top navigation bar, enter
                    your <code>EXCEL-XXXX-XXXX...</code> activation key, and click{' '}
                    <strong>Activate</strong>. The key adds 30 days, 60 days, or extended access
                    directly to your account after the server verifies it.
                  </p>
                </details>

                <details className="faq-item">
                  <summary>How long is the evaluation?</summary>
                  <p>
                    The official hosted deployment gives each verified email account a 30-day
                    evaluation. After it ends, request an activation key through support; there is
                    no in-app checkout or automatic renewal.
                  </p>
                </details>

                <details className="faq-item">
                  <summary>Does activation pay for AI model usage?</summary>
                  <p>
                    No. An activation key controls access to the official hosted deployment. Local
                    deterministic operations need no AI key when available, while BYOK
                    model-provider usage and any provider costs remain separate.
                  </p>
                </details>

                <details className="faq-item">
                  <summary>Does activation include unattended automation?</summary>
                  <p>
                    No. Activation unlocks the browser workspace and its reviewable workflow. It
                    does not create a background workbook queue or an unattended automation service.
                  </p>
                </details>

                <details className="faq-item">
                  <summary>Can I import password-protected or CSV files?</summary>
                  <p>
                    CSV, TSV, and standard Excel (.xlsx) files up to 100,000 rows are supported
                    natively. Password-protected files must have their password removed prior to
                    ingestion.
                  </p>
                </details>

                <details className="faq-item">
                  <summary>Where is my data stored?</summary>
                  <p>
                    Your spreadsheet data remains local in your browser sandbox. Only formula and
                    transformation requests sent to AI models are routed through your chosen
                    provider (OpenRouter / BYOK).
                  </p>
                </details>
              </div>

              {onNavigate && (
                <button
                  type="button"
                  className="support-docs-link"
                  onClick={() => onNavigate('docs')}
                >
                  Explore Full Documentation <ArrowRight size={14} />
                </button>
              )}
            </div>
          </aside>
        </div>
      </main>

      <ResourcePageFooter current="support" />
    </div>
  );
};
