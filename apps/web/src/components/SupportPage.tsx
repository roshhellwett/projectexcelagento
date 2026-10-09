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
  { id: 'license', label: 'License & Key Activation', icon: '🔑' },
  { id: 'bug', label: 'Bug or Calculation Error', icon: '🐛' },
  { id: 'feature', label: 'Feature Request / AI Capability', icon: '💡' },
  { id: 'billing', label: 'Account & Plan Inquiries', icon: '💳' },
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
    <div className="app-container app-page-scroll resource-page-container">
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
            Have questions about activation keys, trial duration, calculation accuracy, or custom
            enterprise requirements? Submit your ticket below and our team will get back to your
            email directly.
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
                      Average response time: <strong>&lt; 24 hours</strong>
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
                  <span className="support-priority-tag">Response SLA: 24h</span>
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
                    Your information is stored securely in compliance with the DPDP Act 2023.
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
                Need urgent assistance or commercial inquiries? Contact our operations inbox
                directly:
              </p>
              <div className="support-email-pill">
                <Mail size={14} />
                <a href="mailto:zenithprojects@icloud.com">zenithprojects@icloud.com</a>
              </div>
              <div className="support-sla-banner">
                <Sparkles size={15} />
                <span>Operating hours: Monday – Saturday, 9:00 AM – 8:00 PM IST</span>
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
                    Go to <strong>Account &amp; Activation</strong> in the top navigation bar, enter
                    your <code>EXCEL-XXXX-XXXX...</code> activation key, and click{' '}
                    <strong>Activate</strong>. The key adds 30 days, 60 days, or extended access
                    directly to your account.
                  </p>
                </details>

                <details className="faq-item">
                  <summary>How many days does the free trial include?</summary>
                  <p>
                    All registered email accounts automatically receive a 30-day free trial. Once
                    the 30 days complete, you can purchase an activation key for 30 days, 60 days,
                    or longer periods to continue using your workspace.
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
