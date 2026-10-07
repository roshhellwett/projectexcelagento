import React from 'react';
import {
  ArrowLeft,
  ArrowUpRight,
  BookOpen,
  Bot,
  FileText,
  Home,
  LifeBuoy,
  Shield,
} from 'lucide-react';

export type ResourcePageId =
  'agents' | 'docs' | 'privacy' | 'terms' | 'missions' | 'usage' | 'support';

const links: Array<{
  id: ResourcePageId | 'landing';
  label: string;
  icon: typeof Home;
  href: string;
}> = [
  { id: 'landing', label: 'Overview', icon: Home, href: '#/landing' },
  { id: 'agents', label: 'Agents', icon: Bot, href: '#/agents' },
  { id: 'docs', label: 'Docs', icon: BookOpen, href: '#/docs' },
  { id: 'support', label: 'Support', icon: LifeBuoy, href: '#/support' },
  { id: 'privacy', label: 'Privacy', icon: Shield, href: '#/privacy' },
  { id: 'terms', label: 'Terms', icon: FileText, href: '#/terms' },
];

const labels: Record<ResourcePageId, string> = {
  agents: 'Agent architecture',
  docs: 'Documentation',
  support: 'Help & support',
  privacy: 'Privacy & data flow',
  terms: 'Terms & governance',
  missions: 'Mission control',
  usage: 'Model & usage',
};

interface ResourcePageChromeProps {
  current: ResourcePageId;
  onBack: () => void;
}

export const ResourcePageChrome: React.FC<ResourcePageChromeProps> = ({ current, onBack }) => (
  <header className="resource-nav">
    <a className="resource-brand" href="#/landing">
      <span className="resource-brand-mark">
        <img src="/excel-agent-logo.svg" alt="" />
      </span>
      <span>
        <strong>
          Excel<span>Agento</span>
        </strong>
        <small>RESOURCE CENTER</small>
      </span>
    </a>
    <nav className="resource-links" aria-label="Resource navigation">
      {links.map(({ id, label, icon: Icon, href }) => (
        <a
          key={id}
          className={id === current ? 'is-active' : ''}
          href={href}
          aria-current={id === current ? 'page' : undefined}
        >
          <Icon size={14} /> {label}
        </a>
      ))}
    </nav>
    <div className="resource-nav-actions">
      <span className="resource-current-label">{labels[current]}</span>
      <button type="button" className="resource-back-button" onClick={onBack}>
        <ArrowLeft size={14} /> Workspace
      </button>
    </div>
  </header>
);

export function ResourcePageFooter({ current }: { current: ResourcePageId }) {
  return (
    <footer className="resource-footer">
      <span>
        <strong>ExcelAgento</strong> · open-source spreadsheet workspace
      </span>
      <span className="resource-footer-current">
        Viewing {labels[current]} <ArrowUpRight size={13} />
      </span>
    </footer>
  );
}
