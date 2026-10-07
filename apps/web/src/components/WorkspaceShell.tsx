import React, { useEffect, useRef, useState } from 'react';
import { motion, useReducedMotion } from 'framer-motion';
import {
  Table2,
  ChartNoAxesCombined,
  Shapes,
  PanelRightClose,
  PanelRightOpen,
  Sparkles,
  ArrowUpRight,
  GripVertical,
} from 'lucide-react';

export type StudioView = 'sheet' | 'insights' | 'workflows';

interface WorkspaceShellProps {
  view: StudioView;
  onViewChange: (view: StudioView) => void;
  sheetName: string;
  isProcessing: boolean;
  agent: React.ReactNode;
  revealAgentRevision: number;
  ribbon?: React.ReactNode;
  children: React.ReactNode;
}

const VIEWS = [
  { id: 'sheet', label: 'Sheet', icon: Table2 },
  { id: 'insights', label: 'Insights', icon: ChartNoAxesCombined },
  { id: 'workflows', label: 'Workflows', icon: Shapes },
] as const;

export function WorkspaceShell({
  view,
  onViewChange,
  sheetName,
  isProcessing,
  agent,
  revealAgentRevision,
  ribbon,
  children,
}: WorkspaceShellProps) {
  const [agentVisible, setAgentVisible] = useState(true);
  const [mobilePanel, setMobilePanel] = useState<'data' | 'agent'>('data');
  const [agentWidth, setAgentWidth] = useState(400);
  const reducedMotion = useReducedMotion();
  const layoutRef = useRef<HTMLDivElement>(null);
  const dragRef = useRef<{ x: number; width: number } | null>(null);

  useEffect(() => {
    if (revealAgentRevision === 0) return;
    setAgentVisible(true);
    setMobilePanel('agent');
  }, [revealAgentRevision]);

  useEffect(() => {
    setMobilePanel('data');
  }, [view]);

  return (
    <main
      className="studio-workspace"
      data-mobile-panel={mobilePanel}
      data-agent-visible={agentVisible}
    >
      <nav className="studio-rail" aria-label="Workspace views">
        <div className="studio-rail-top">
          {VIEWS.map(({ id, label, icon: Icon }) => (
            <button
              key={id}
              type="button"
              className={`studio-rail-item ${view === id ? 'active' : ''}`}
              aria-pressed={view === id}
              onClick={() => {
                onViewChange(id);
                setMobilePanel('data');
              }}
              title={label}
              data-sound="toggle"
            >
              {view === id && (
                <motion.span
                  className="studio-rail-active"
                  layoutId="studio-navigation"
                  transition={
                    reducedMotion
                      ? { duration: 0 }
                      : { type: 'spring', stiffness: 360, damping: 32 }
                  }
                />
              )}
              <Icon size={20} strokeWidth={1.7} />
              <span>{label}</span>
            </button>
          ))}
        </div>
        <div className="studio-rail-bottom">
          <span className="studio-rail-monogram" aria-hidden="true">
            EA
          </span>
          <span className="studio-rail-caption">
            Built for
            <br />
            your next idea.
          </span>
        </div>
      </nav>
      <div
        className="studio-layout"
        ref={layoutRef}
        style={{ '--agent-width': `${agentWidth}px` } as React.CSSProperties}
      >
        <div className="studio-mobile-tabs" role="group" aria-label="Workspace panel">
          <button
            type="button"
            aria-pressed={mobilePanel === 'data'}
            onClick={() => setMobilePanel('data')}
            data-sound="toggle"
          >
            <Table2 size={15} />
            Workbook
          </button>
          <button
            type="button"
            aria-pressed={mobilePanel === 'agent'}
            onClick={() => {
              setMobilePanel('agent');
              setAgentVisible(true);
            }}
            data-sound="toggle"
          >
            <Sparkles size={15} />
            Agent {isProcessing && <span className="studio-status-dot busy" />}
          </button>
        </div>
        <section className="studio-data-area" aria-label="Workbook workspace">
          <div className="studio-stage-heading">
            <div>
              <span className="studio-eyebrow">YOUR WORKING CANVAS</span>
              <h1>
                {view === 'sheet'
                  ? sheetName
                  : view === 'insights'
                    ? 'A clearer picture.'
                    : 'What will you create?'}
              </h1>
            </div>
            <button
              type="button"
              className="btn btn-ghost btn-sm studio-agent-toggle"
              onClick={() => setAgentVisible(!agentVisible)}
              aria-expanded={agentVisible}
              aria-controls="studio-agent-area"
              title={agentVisible ? 'Focus on workbook' : 'Show agent'}
              data-sound="toggle"
            >
              {agentVisible ? <PanelRightClose size={17} /> : <PanelRightOpen size={17} />}
              <span>{agentVisible ? 'Focus mode' : 'Show agent'}</span>
            </button>
          </div>
          {ribbon}
          <div className="studio-stage-content">{children}</div>
        </section>
        {agentVisible && (
          <div
            className="studio-panel-splitter"
            role="separator"
            tabIndex={0}
            aria-label="Resize agent panel"
            aria-orientation="vertical"
            aria-valuenow={agentWidth}
            aria-valuemin={340}
            aria-valuemax={560}
            onKeyDown={(event) => {
              if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return;
              event.preventDefault();
              event.stopPropagation();
              setAgentWidth((width) =>
                Math.min(560, Math.max(340, width + (event.key === 'ArrowLeft' ? 20 : -20))),
              );
            }}
            onPointerDown={(event) => {
              if (event.button !== 0) return;
              event.preventDefault();
              dragRef.current = { x: event.clientX, width: agentWidth };
              event.currentTarget.setPointerCapture(event.pointerId);
            }}
            onPointerMove={(event) => {
              if (!dragRef.current) return;
              const available = (layoutRef.current?.clientWidth ?? 1000) - 380;
              setAgentWidth(
                Math.max(
                  340,
                  Math.min(
                    560,
                    available,
                    dragRef.current.width + dragRef.current.x - event.clientX,
                  ),
                ),
              );
            }}
            onPointerUp={() => {
              dragRef.current = null;
            }}
            onLostPointerCapture={() => {
              dragRef.current = null;
            }}
          >
            <GripVertical size={13} />
          </div>
        )}
        <section
          id="studio-agent-area"
          className="studio-agent-area"
          aria-label="Agent workspace"
          hidden={!agentVisible}
        >
          {agent}
        </section>
        {!agentVisible && (
          <button
            type="button"
            className="studio-agent-reopen"
            onClick={() => setAgentVisible(true)}
            data-sound="toggle"
          >
            <Sparkles size={19} />
            <span>Agent</span>
            <ArrowUpRight size={14} />
          </button>
        )}
      </div>
    </main>
  );
}
