import { useRef, useState } from 'react';
import {
  ArrowLeft,
  ClipboardCheck,
  Clock3,
  Play,
  Trash2,
  ShieldCheck,
  Download,
  RotateCcw,
  Square,
} from 'lucide-react';
import type { MissionRecord } from '../lib/missions.js';
import type { MissionStorageStatus } from '../lib/use-mission-ledger.js';
import { useDialogA11y } from '../lib/use-dialog-a11y.js';
import { ResourcePageChrome, ResourcePageFooter } from './ResourcePageChrome.js';

interface MissionControlPageProps {
  missions: MissionRecord[];
  storageStatus: MissionStorageStatus;
  storageError: string;
  unsavedIds: Set<string>;
  busy: boolean;
  activeMissionId: string | null;
  onStop: (mission: MissionRecord) => void;
  onBack: () => void;
  onResume: (mission: MissionRecord) => void;
  onReplan: (mission: MissionRecord) => void;
  onDelete: (mission: MissionRecord) => Promise<void>;
  onClear: () => Promise<void>;
  onRetryStorage: () => void;
  onDownload: (mission: MissionRecord) => void;
}
const statusCopy: Record<MissionRecord['status'], string> = {
  planning: 'Planning',
  executing: 'Verifying execution',
  prepared: 'Ready for review',
  analyzed: 'Analysis complete',
  applied: 'Completed',
  undone: 'Reverted',
  failed: 'Failed safely',
  stale: 'Needs re-check',
  cancelled: 'Cancelled',
  interrupted: 'Interrupted',
};
const time = (value: number) => new Date(value).toLocaleString();

export function MissionControlPage({
  missions,
  storageStatus,
  storageError,
  unsavedIds,
  busy,
  activeMissionId,
  onStop,
  onBack,
  onResume,
  onReplan,
  onDelete,
  onClear,
  onRetryStorage,
  onDownload,
}: MissionControlPageProps) {
  const [filter, setFilter] = useState('all');
  const [deleting, setDeleting] = useState<MissionRecord | 'all' | null>(null);
  const [deleteError, setDeleteError] = useState('');
  const [isDeleting, setIsDeleting] = useState(false);
  const dialog = useRef<HTMLDivElement>(null);
  useDialogA11y(deleting !== null, dialog, () => {
    if (!isDeleting) setDeleting(null);
  });
  const visible = missions.filter(
    (mission) =>
      filter === 'all' ||
      (filter === 'review'
        ? mission.status === 'prepared'
        : filter === 'completed'
          ? ['analyzed', 'applied', 'undone'].includes(mission.status)
          : ['failed', 'stale', 'cancelled', 'interrupted'].includes(mission.status)),
  );
  const confirmDelete = async () => {
    if (!deleting) return;
    setIsDeleting(true);
    setDeleteError('');
    try {
      if (deleting === 'all') await onClear();
      else await onDelete(deleting);
      setDeleting(null);
    } catch {
      setDeleteError(
        'Deletion did not commit. Stored mission data may still exist. Retry or keep the history.',
      );
    } finally {
      setIsDeleting(false);
    }
  };
  return (
    <div className="mission-page resource-page resource-page-missions" data-testid="mission-page">
      <ResourcePageChrome current="missions" onBack={onBack} />
      <div className="mission-topbar">
        <button type="button" className="btn btn-secondary btn-sm" onClick={onBack}>
          <ArrowLeft size={14} /> Back to workspace
        </button>
        <span className="zenith-brand-badge">
          <ClipboardCheck size={13} /> Mission control · local-first
        </span>
      </div>
      <header className="mission-hero">
        <div className="mission-hero-badge">
          <ShieldCheck size={14} /> WORKBOOK MISSIONS
        </div>
        <h1>Work that survives the chat.</h1>
        <p>
          Goals, reviewed plans, evidence and execution receipts in one place. Resume a prepared
          task only against its exact workbook content—or replan it for the workbook you have now.
        </p>
        <div className="mission-hero-meta">
          <span>
            <strong>{missions.length}</strong> mission{missions.length === 1 ? '' : 's'}
          </span>
          <span>
            <strong>{missions.filter((mission) => mission.status === 'prepared').length}</strong>{' '}
            awaiting review
          </span>
          <span role="status" aria-label="Mission storage status">
            {storageStatus === 'checking'
              ? 'Checking browser storage…'
              : storageStatus === 'saving'
                ? 'Saving mission history…'
                : storageStatus === 'unavailable'
                  ? 'History is not fully saved'
                  : 'Mission history saved locally'}
          </span>
        </div>
      </header>
      {storageStatus === 'unavailable' && (
        <section className="mission-storage-warning" role="alert">
          <div>
            <strong>Keep working, but don’t rely on refresh recovery.</strong>
            <p>
              {storageError} Unsaved missions are kept in memory. Previous committed history remains
              in storage.
            </p>
          </div>
          <button type="button" className="btn btn-secondary btn-sm" onClick={onRetryStorage}>
            Retry mission storage
          </button>
        </section>
      )}
      <section className="mission-toolbar" aria-label="Mission controls">
        <div>
          <span className="studio-eyebrow">MISSION LOG</span>
          <p>
            Latest 50 tasks. Requests, evidence and operation arguments may contain workbook data.
          </p>
        </div>
        <button
          type="button"
          className="btn btn-ghost btn-sm"
          disabled={(!missions.length && storageStatus !== 'unavailable') || busy}
          onClick={() => {
            setDeleteError('');
            setDeleting('all');
          }}
        >
          <Trash2 size={13} /> Clear mission history
        </button>
      </section>
      <div className="mission-filters" role="group" aria-label="Filter missions">
        {[
          ['all', 'All missions'],
          ['review', 'Awaiting review'],
          ['completed', 'Completed'],
          ['attention', 'Needs attention'],
        ].map(([id, label]) => (
          <button
            type="button"
            className="btn btn-ghost btn-sm"
            key={id}
            aria-pressed={filter === id}
            onClick={() => setFilter(id!)}
          >
            {label}
          </button>
        ))}
      </div>
      {!visible.length ? (
        <div className="mission-empty">
          <Clock3 size={22} />
          <h2>{missions.length ? 'No missions in this view' : 'Your next task starts here'}</h2>
          <p>
            Ask for an analysis or transformation in the workspace. Nothing runs automatically after
            refresh; prepared work returns for your review.
          </p>
          <button type="button" className="btn btn-primary" onClick={onBack}>
            Open workspace
          </button>
        </div>
      ) : (
        <div className="mission-list">
          {visible.map((mission) => (
            <article
              className={`mission-card is-${mission.status}`}
              key={mission.id}
              aria-label={`Mission: ${mission.title}`}
            >
              <div className="mission-card-head">
                <div>
                  <span className="mission-kind">
                    {mission.kind === 'plan'
                      ? 'MULTI-STEP PLAN'
                      : mission.kind === 'analysis'
                        ? 'ANALYST TASK'
                        : 'WORKBOOK ACTION'}
                  </span>
                  <h2>{mission.title}</h2>
                </div>
                <span className="mission-status">{statusCopy[mission.status]}</span>
              </div>
              <p className="mission-request">{mission.request}</p>
              <div className="mission-context">
                <span>{mission.workbook.fileName}</span>
                <span>{mission.workbook.sheetName}</span>
                <time dateTime={new Date(mission.updatedAt).toISOString()}>
                  {time(mission.updatedAt)}
                </time>
                {unsavedIds.has(mission.id) && (
                  <span className="mission-unsaved">Not saved yet</span>
                )}
              </div>
              {mission.plan && (
                <ol className="mission-step-list">
                  {mission.plan.steps.map((step, index) => (
                    <li key={`${step.id}-${index}`}>
                      <span>{index + 1}</span>
                      <div>
                        <strong>{step.description}</strong>
                        <small>
                          {step.operation}
                          {mission.receipt?.operations[index]
                            ? ` · ${mission.receipt.operations[index]!.affectedCells.toLocaleString()} affected cells`
                            : ' · not applied'}
                        </small>
                      </div>
                    </li>
                  ))}
                </ol>
              )}
              {mission.receipt && (
                <div className="mission-receipt">
                  <strong>
                    {mission.receipt.status === 'undone'
                      ? 'Commit reverted'
                      : mission.receipt.status === 'failed'
                        ? 'Nothing committed'
                        : 'Engine-verified commit'}
                  </strong>
                  <span>
                    {mission.receipt.operations
                      .reduce((total, operation) => total + operation.affectedCells, 0)
                      .toLocaleString()}{' '}
                    cell changes across {mission.receipt.operations.length} operation
                    {mission.receipt.operations.length === 1 ? '' : 's'}
                  </span>
                  <p>{mission.receipt.operations.flatMap((op) => op.warnings).join(' ')}</p>
                </div>
              )}
              {mission.evidence?.length ? (
                <details className="mission-evidence">
                  <summary>
                    <ShieldCheck size={13} /> {mission.evidence.length} evidence source
                    {mission.evidence.length === 1 ? '' : 's'}
                  </summary>
                  {mission.evidence.map((item, index) => (
                    <div className="mission-evidence-item" key={`${item.id}-${index}`}>
                      <strong>{item.title}</strong>
                      <code>{item.source}</code>
                      <dl>
                        {item.facts.map((fact) => (
                          <div key={fact.label}>
                            <dt>{fact.label}</dt>
                            <dd>{fact.value}</dd>
                          </div>
                        ))}
                      </dl>
                      {item.note && <p>{item.note}</p>}
                    </div>
                  ))}
                </details>
              ) : null}
              {(mission.status === 'planning' || mission.status === 'executing') && (
                <p className="mission-progress" role="status">
                  {mission.stage ?? 'Inspecting workbook…'}
                </p>
              )}
              {mission.answer && (
                <details className="mission-answer">
                  <summary>Read saved result</summary>
                  <p>{mission.answer}</p>
                  <small>
                    Saved interpretation from the inspected workbook; it is not an independently
                    verified conclusion.
                  </small>
                </details>
              )}
              {mission.error && <p className="mission-error">{mission.error}</p>}
              {!mission.workbook.signature && mission.status === 'prepared' && (
                <p className="mission-error">
                  Secure content verification was unavailable; use the current chat preview or
                  replan, not refresh resume.
                </p>
              )}
              <div className="mission-actions">
                {(mission.status === 'planning' || mission.status === 'executing') &&
                  activeMissionId === mission.id && (
                    <button
                      type="button"
                      className="btn btn-secondary btn-sm"
                      onClick={() => onStop(mission)}
                    >
                      <Square size={13} /> Stop mission
                    </button>
                  )}
                {mission.status === 'prepared' && mission.workbook.signature && (
                  <button
                    type="button"
                    className="btn btn-primary btn-sm"
                    disabled={busy || isDeleting}
                    onClick={() => onResume(mission)}
                  >
                    <Play size={13} /> Resume review
                  </button>
                )}
                {mission.status !== 'planning' && mission.status !== 'executing' && (
                  <button
                    type="button"
                    className="btn btn-secondary btn-sm"
                    disabled={busy || isDeleting}
                    onClick={() => onReplan(mission)}
                  >
                    <RotateCcw size={13} /> Replan request
                  </button>
                )}
                <button
                  type="button"
                  className="btn btn-ghost btn-sm"
                  onClick={() => onDownload(mission)}
                  title="Download contains the mission request, evidence and operation arguments"
                >
                  <Download size={13} /> Download record
                </button>
                <button
                  type="button"
                  className="btn btn-ghost btn-sm"
                  disabled={busy || isDeleting}
                  onClick={() => {
                    setDeleteError('');
                    setDeleting(mission);
                  }}
                >
                  <Trash2 size={13} /> Remove
                </button>
              </div>
            </article>
          ))}
        </div>
      )}
      <p className="mission-footnote">
        Browser-local records are not encrypted by the app. Prepared plans survive refresh; active
        model calls do not. Atomic execution prevents partial workbook commits, so recovery resumes
        review, not half-applied mutations.
      </p>
      {deleting && (
        <div className="workspace-decision-overlay">
          <div
            ref={dialog}
            className="workspace-decision-card"
            role="dialog"
            aria-modal="true"
            aria-labelledby="mission-delete-title"
            data-dialog-open="true"
          >
            <h2 id="mission-delete-title">
              {deleting === 'all' ? 'Delete local mission history?' : 'Remove this mission?'}
            </h2>
            <p>
              Delete saved requests, operation arguments, evidence and receipts{' '}
              {deleting === 'all' ? 'for all listed missions' : 'for this task'}. This does not undo
              workbook changes, clear workbook checkpoints or delete provider records.
            </p>
            {deleteError && <p role="alert">{deleteError}</p>}
            <div className="mission-actions">
              <button
                type="button"
                className="btn btn-secondary"
                disabled={isDeleting}
                onClick={() => setDeleting(null)}
              >
                Keep history
              </button>
              <button
                type="button"
                className="btn btn-danger"
                disabled={isDeleting}
                onClick={() => void confirmDelete()}
              >
                {isDeleting ? 'Deleting…' : 'Delete mission data'}
              </button>
            </div>
          </div>
        </div>
      )}
      <ResourcePageFooter current="missions" />
    </div>
  );
}
