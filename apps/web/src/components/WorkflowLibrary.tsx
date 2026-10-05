import { useMemo, useState } from 'react';
import { motion, useReducedMotion } from 'framer-motion';
import { ArrowUpRight, Search, Sparkles, Copy, ScanSearch, ChartNoAxesCombined, ListFilter, ArrowDownWideNarrow, CalendarDays, Table2, ShieldCheck } from 'lucide-react';
import type { CompactColumnProfile, ToolDescriptor } from '@excel-agent/agent';
import { workspaceWorkflows, type WorkspaceWorkflow } from '../lib/workflows.js';

export const WORKFLOW_ICONS = {
  sparkles: Sparkles, duplicates: Copy, missing: ScanSearch, statistics: ChartNoAxesCombined,
  outliers: ListFilter, sort: ArrowDownWideNarrow, dates: CalendarDays, overview: Table2,
};

interface WorkflowLibraryProps {
  profiles: CompactColumnProfile[];
  tools: ToolDescriptor[];
  onRun: (prompt: string) => void;
  onDraft: (prompt: string) => void;
  isProcessing: boolean;
}

export function WorkflowLibrary({ profiles, tools, onRun, onDraft, isProcessing }: WorkflowLibraryProps) {
  const [category, setCategory] = useState('All');
  const [search, setSearch] = useState('');
  const reducedMotion = useReducedMotion();
  const workflows = useMemo(() => workspaceWorkflows(profiles), [profiles]);
  const filtered = workflows.filter((workflow) => (category === 'All' || workflow.category === category) && `${workflow.title} ${workflow.description}`.toLowerCase().includes(search.toLowerCase()));
  const filteredTools = tools.filter((tool) => `${tool.name} ${tool.description}`.toLowerCase().includes(search.toLowerCase()));
  return (
    <div className="studio-scroll-page workflow-library">
      <div className="workflow-hero"><span className="studio-eyebrow">LESS REPETITION. MORE POSSIBILITY.</span><h2>Your next task,<br /><span>already one step closer.</span></h2><p>Start with a workflow built for this sheet. The agent inspects the data, proposes a change, and lets you review it.</p><span className="workflow-hero-mark" aria-hidden="true"><ShapesIllustration /></span></div>
      <div className="workflow-library-controls">
        <div className="studio-segmented" role="group" aria-label="Workflow category">{['All', 'Clean', 'Analyze', 'Organize'].map((item) => <button type="button" key={item} aria-pressed={category === item} onClick={() => setCategory(item)}>{item}</button>)}</div>
        <label className="studio-filter-input"><Search size={15} /><span className="sr-only">Search workflows</span><input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Find a workflow…" /></label>
      </div>
      <div className="workflow-card-grid">
        {filtered.map((workflow, index) => <WorkflowCard key={workflow.id} workflow={workflow} index={index} reducedMotion={Boolean(reducedMotion)} onRun={onRun} disabled={isProcessing} />)}
        {filtered.length === 0 && <p className="studio-empty-message">No matching workflows. Try another search or browse the engine catalog below.</p>}
      </div>
      <div className="studio-section-heading"><div><span className="studio-eyebrow">THE ENGINE BEHIND THE AGENT</span><h3>{tools.length} operations. One workspace.</h3></div><span className="studio-soft-badge"><ShieldCheck size={13} />Validated operations</span></div>
      <div className="workflow-engine-catalog">
        {filteredTools.map((tool) => <details key={tool.name} className="workflow-engine-item"><summary><span>{tool.name.replace(/_/g, ' ')}</span><span className="studio-catalog-category">{tool.category}</span></summary><p>{tool.description}</p><button type="button" className="btn btn-secondary btn-sm" disabled={isProcessing} onClick={() => onDraft(`Help me use ${tool.name} on the current sheet. Ask for the required columns and settings, then preview the operation.`)}>Plan with agent <ArrowUpRight size={13} /></button></details>)}
      </div>
    </div>
  );
}

function WorkflowCard({ workflow, index, reducedMotion, onRun, disabled }: { workflow: WorkspaceWorkflow; index: number; reducedMotion: boolean; onRun: (prompt: string) => void; disabled: boolean }) {
  const Icon = WORKFLOW_ICONS[workflow.icon];
  return <motion.button type="button" className="workflow-card" disabled={disabled} onClick={() => onRun(workflow.prompt)}
    initial={reducedMotion ? false : { opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} transition={{ duration: 0.25, delay: Math.min(index * 0.025, 0.15) }}>
    <span className={`workflow-icon workflow-icon-${workflow.category.toLowerCase()}`}><Icon size={20} strokeWidth={1.6} /></span><span className="workflow-card-category">{workflow.category}</span><strong>{workflow.title}</strong><span className="workflow-card-description">{workflow.description}</span><span className="workflow-card-action">Start workflow <ArrowUpRight size={15} /></span>
  </motion.button>;
}

function ShapesIllustration() {
  return <svg viewBox="0 0 160 150" fill="none"><rect x="15" y="42" width="80" height="80" rx="18" transform="rotate(-12 15 42)" stroke="currentColor" strokeWidth="1.5" /><rect x="60" y="20" width="80" height="80" rx="18" transform="rotate(12 60 20)" fill="currentColor" fillOpacity=".08" stroke="currentColor" strokeWidth="1.5" /><path d="m65 78 13 13 26-28" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" /><circle cx="122" cy="123" r="12" stroke="currentColor" strokeWidth="1.5" /><path d="M122 117v12m-6-6h12" stroke="currentColor" strokeWidth="1.5" /></svg>;
}
