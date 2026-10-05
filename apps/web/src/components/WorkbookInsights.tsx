import React, { useMemo, useState } from 'react';
import { ArrowUpRight, ChartNoAxesCombined, CircleCheck, CircleDashed, Hash, Sparkles, Type, CalendarDays, ScanSearch } from 'lucide-react';
import type { CompactColumnProfile, SheetAudit } from '@excel-agent/agent';

interface WorkbookInsightsProps {
  profiles: CompactColumnProfile[];
  audit: SheetAudit;
  onRun: (prompt: string) => void;
  isProcessing: boolean;
}

export function WorkbookInsights({ profiles, audit, onRun, isProcessing }: WorkbookInsightsProps) {
  const numericColumns = useMemo(() => profiles.filter((column) => column.isNumeric), [profiles]);
  const [xColumn, setXColumn] = useState('');
  const [yColumn, setYColumn] = useState('');
  const x = numericColumns.find((column) => column.letter === xColumn) ?? numericColumns[0];
  const y = numericColumns.find((column) => column.letter === yColumn) ?? numericColumns[1];
  const rows = Math.max(0, audit.totalRows - 1);
  const totalCells = rows * profiles.length;
  const populated = profiles.reduce((sum, column) => sum + column.nonBlankCount, 0);
  const missing = Math.max(0, totalCells - populated);
  const completeness = totalCells > 0 ? (populated / totalCells) * 100 : null;
  return (
    <div className="studio-scroll-page workbook-insights">
      <div className="insights-intro"><p>Your data, before the next decision.</p><span className="studio-soft-badge"><CircleCheck size={13} />Computed in your browser</span></div>
      <div className="insights-metrics">
        <Metric label="DATA RECORDS" value={rows.toLocaleString()} detail="Excluding the header row" icon={<Hash size={17} />} />
        <Metric label="DATA COMPLETENESS" value={completeness === null ? '—' : `${completeness.toFixed(1)}%`} detail="Nonblank cells across this sheet" icon={<CircleCheck size={17} />} />
        <Metric label="MISSING VALUES" value={missing.toLocaleString()} detail="Empty cells, including ragged rows" icon={<CircleDashed size={17} />} />
      </div>
      <div className="insights-two-column">
        <section className="insights-card">
          <div className="studio-section-heading"><h3>A little attention goes a long way.</h3><ScanSearch size={18} /></div>
          <p className="studio-section-description">The workbook audit found these opportunities.</p>
          {audit.findings.length > 0 ? <ul className="insights-findings">{audit.findings.slice(0, 5).map((finding, index) => <li key={index}><span className="insights-finding-dot" />{finding}</li>)}</ul> : <p className="insights-clean-note"><CircleCheck size={18} />No issues flagged by the current audit.</p>}
          <button type="button" className="btn btn-secondary btn-sm" disabled={isProcessing} onClick={() => onRun('audit missing values and empty cells')}>Inspect missing values <ArrowUpRight size={14} /></button>
        </section>
        <section className="insights-card insights-relationship">
          <div className="studio-section-heading"><h3>Find the connection.</h3><ChartNoAxesCombined size={18} /></div>
          <p className="studio-section-description">Compare two numeric columns with complete-pair correlation or a simple linear regression.</p>
          {numericColumns.length >= 2 ? <>
            <div className="insights-column-pickers">
              <label>Predictor (X)<select value={x?.letter ?? ''} onChange={(event) => setXColumn(event.target.value)}>{numericColumns.map((column) => <option key={column.letter} value={column.letter}>{column.rawName} ({column.letter})</option>)}</select></label>
              <label>Response (Y)<select value={y?.letter ?? ''} onChange={(event) => setYColumn(event.target.value)}>{numericColumns.map((column) => <option key={column.letter} value={column.letter}>{column.rawName} ({column.letter})</option>)}</select></label>
            </div>
            <div className="insights-relationship-actions"><button type="button" className="btn btn-primary btn-sm" disabled={isProcessing || x?.letter === y?.letter} onClick={() => onRun(`correlation between column ${x!.letter} and column ${y!.letter}`)}>Correlation <ArrowUpRight size={14} /></button><button type="button" className="btn btn-secondary btn-sm" disabled={isProcessing || x?.letter === y?.letter} onClick={() => onRun(`linear regression of column ${y!.letter} on column ${x!.letter}`)}>Linear regression</button></div>
            {x?.letter === y?.letter && <p className="studio-field-hint">Choose two different columns to compare.</p>}
          </> : <p className="studio-empty-message">This analysis needs at least two numeric columns. You can still inspect individual columns below.</p>}
        </section>
      </div>
      <div className="studio-section-heading"><div><span className="studio-eyebrow">A FIELD GUIDE TO YOUR DATA</span><h3>Every column tells a story.</h3></div><span className="studio-soft-badge">{profiles.length} columns</span></div>
      <div className="insights-column-list">
        {profiles.map((column) => {
          const blank = Math.max(0, rows - column.nonBlankCount);
          const filledPercent = rows > 0 ? (column.nonBlankCount / rows) * 100 : 0;
          const Icon = column.isNumeric ? Hash : column.isDate ? CalendarDays : Type;
          return <div key={column.letter} className="insights-column-row"><span className="insights-column-letter">{column.letter}</span><div className="insights-column-name"><strong>{column.rawName}</strong><span><Icon size={11} />{column.isNumeric ? 'Numeric' : column.isDate ? 'Date' : 'Text'} · {column.distinctCountIsLowerBound ? 'at least ' : ''}{column.distinctCount} unique</span></div><div className="insights-completeness"><div className="insights-completeness-track"><span style={{ width: `${filledPercent}%` }} /></div><span>{blank === 0 ? 'Complete' : `${blank} missing`}</span></div><button type="button" className="btn btn-ghost btn-sm" disabled={isProcessing} onClick={() => onRun(`descriptive statistics for column ${column.letter}`)} aria-label={`Analyze ${column.rawName} column ${column.letter}`}><Sparkles size={14} /><span>Analyze</span></button></div>;
        })}
        {profiles.length === 0 && <p className="studio-empty-message">Upload a workbook to start exploring its columns.</p>}
      </div>
    </div>
  );
}

function Metric({ label, value, detail, icon }: { label: string; value: string; detail: string; icon: React.ReactNode }) {
  return <div className="insights-metric"><div><span className="studio-eyebrow">{label}</span>{icon}</div><strong>{value}</strong><span>{detail}</span></div>;
}
