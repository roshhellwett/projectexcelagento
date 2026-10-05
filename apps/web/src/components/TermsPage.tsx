import React from 'react';
import { ArrowLeft, Scale, Sparkles, ExternalLink, ShieldAlert } from 'lucide-react';

interface TermsPageProps {
  onBack: () => void;
}

export const TermsPage: React.FC<TermsPageProps> = ({ onBack }) => {
  return (
    <div className="legal-page" data-testid="terms-page">
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
          <Scale size={14} />
          <span>OPEN SOURCE GOVERNANCE &amp; STATUTORY COMPLIANCE</span>
        </div>
        <h1 className="legal-title">Terms of Service &amp; Open Source Governance</h1>
        <p className="legal-effective">
          Governed under the Indian Information Technology Act, 2000 &bull; Intermediary Guidelines
          Rules, 2021 &bull; Open-Source MIT License.
        </p>
        <p className="legal-updated">
          Effective Date: October 2026 &bull; Published by Zenith Open Source Projects
        </p>
      </header>

      <article className="legal-content">
        <section className="legal-section">
          <h2>1. Acceptance of Terms &amp; Open Source License</h2>
          <p>
            Welcome to <strong>ExcelAgento</strong>, an open-source autonomous spreadsheet
            engineering platform developed and maintained under the banner of{' '}
            <strong>Zenith Open Source Projects</strong> (accessible at{' '}
            <a
              href="https://zenithopensourceprojects.vercel.app/os"
              target="_blank"
              rel="noopener noreferrer"
            >
              https://zenithopensourceprojects.vercel.app/os
            </a>
            ).
          </p>
          <p>
            By accessing, loading, or deploying ExcelAgento, you agree to be bound by these Terms of
            Service and the underlying <strong>MIT Open Source License</strong>. If you do not agree
            to these terms, you must refrain from using the software.
          </p>
        </section>

        <section className="legal-section">
          <h2>2. Permitted Use &amp; Due Diligence (Rule 3(1)(b) IT Rules 2021)</h2>
          <p>
            In compliance with Rule 3(1)(b) of the Information Technology (Intermediary Guidelines
            and Digital Media Ethics Code) Rules, 2021, you agree not to use ExcelAgento to host,
            display, upload, modify, publish, transmit, store, update, or share any information
            that:
          </p>
          <ul>
            <li>Belongs to another person and to which you do not have any right;</li>
            <li>
              Is obscene, pornographic, defamatory, libelous, or racially or ethnically
              objectionable;
            </li>
            <li>Infringes any patent, trademark, copyright, or other proprietary rights;</li>
            <li>Violates any law for the time being in force in the Republic of India;</li>
            <li>
              Deceives or misleads the addressee about the origin of the message or knowingly
              communicates any misinformation;
            </li>
            <li>
              Contains software virus or any other computer code designed to interrupt, destroy, or
              limit the functionality of any computer resource.
            </li>
          </ul>
        </section>

        <section className="legal-section">
          <h2>3. Mathematical &amp; Financial Calculations: Verification Requirement</h2>
          <div className="legal-callout">
            <ShieldAlert size={16} className="callout-icon text-emerald" />
            <div>
              <strong>Professional Due Diligence Notice:</strong> While ExcelAgento employs
              deterministic formula calculation engines and rigorous multi-agent verification,
              artificial intelligence outputs and synthesized formulas must be independently audited
              prior to submission in formal audits, tax filings, legal proceedings, or regulatory
              disclosures.
            </div>
          </div>
          <p>
            ExcelAgento and Zenith Open Source Projects do not provide certified chartered
            accountancy, legal, tax, or investment advice. The user retains full responsibility for
            auditing numerical calculations, CAGR figures, margin reports, and spreadsheet formulas.
          </p>
        </section>

        <section className="legal-section">
          <h2>4. Client-Side Local Execution &amp; Data Sovereignty</h2>
          <p>
            ExcelAgento runs directly on the user's client hardware. Zenith Open Source Projects
            does not host your spreadsheet data, does not intercept your workbook mutations, and has
            no access to files processed within your browser session.
          </p>
        </section>

        <section className="legal-section">
          <h2>5. Disclaimer of Warranties (MIT License &amp; Section 43A IT Act)</h2>
          <p>
            THE SOFTWARE IS PROVIDED &ldquo;AS IS&rdquo;, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
            IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A
            PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS, ZENITH OPEN
            SOURCE PROJECTS, OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
            LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR
            IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
          </p>
        </section>

        <section className="legal-section">
          <h2>6. Limitation of Liability</h2>
          <p>
            To the maximum extent permitted under applicable Indian law, Zenith Open Source
            Projects, its contributors, and developers shall not be liable for any indirect,
            incidental, special, consequential, or punitive damages, including loss of profits, data
            corruption, or business interruption, arising from the use or inability to use this
            software.
          </p>
        </section>

        <section className="legal-section">
          <h2>7. Governing Law &amp; Dispute Resolution</h2>
          <p>
            These Terms of Service shall be governed by and construed in accordance with the laws of
            the <strong>Republic of India</strong>, without regard to its conflict of law
            principles.
          </p>
          <p>
            Any disputes arising out of or related to these Terms shall be subject to the exclusive
            jurisdiction of the competent civil courts situated in{' '}
            <strong>Bengaluru, Karnataka, India</strong> or <strong>New Delhi, India</strong>.
          </p>
        </section>

        <section className="legal-section">
          <h2>8. Contact &amp; Governance</h2>
          <p>For questions regarding these Terms or open-source licensing:</p>
          <p>
            <strong>Zenith Open Source Projects Hub:</strong>{' '}
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

      <footer className="legal-footer">
        <button className="btn btn-primary btn-sm" onClick={onBack}>
          Return to Excel Workspace
        </button>
        <p className="legal-copyright">
          &copy; 2026 Zenith Open Source Projects. Licensed under MIT.
        </p>
      </footer>
    </div>
  );
};
