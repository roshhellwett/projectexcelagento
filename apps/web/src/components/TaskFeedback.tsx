import { useId, useState, type FormEvent } from 'react';
import {
  FEEDBACK_CATEGORY_LABELS,
  getFeedbackStorageStatus,
  type FeedbackCategory,
  type FeedbackInput,
  type FeedbackRating,
  type FeedbackResult,
  type FeedbackStorageStatus,
} from '../lib/feedback.js';

interface TaskFeedbackProps {
  result: FeedbackResult;
  defaultCategory?: FeedbackCategory;
  onSubmit: (feedback: FeedbackInput) => void;
  compact?: boolean;
}

const ratingLabels: Record<FeedbackResult, Array<{ value: FeedbackRating; label: string }>> = {
  completed: [
    { value: 'helpful', label: 'Worked well' },
    { value: 'mixed', label: 'Partly worked' },
    { value: 'not_helpful', label: 'Did not help' },
  ],
  blocked: [
    { value: 'helpful', label: 'Block made sense' },
    { value: 'mixed', label: 'Not sure' },
    { value: 'not_helpful', label: 'I needed this task' },
  ],
};

const categoryOptions = Object.entries(FEEDBACK_CATEGORY_LABELS) as Array<
  [FeedbackCategory, string]
>;

/** A deliberately context-free feedback prompt: no workbook or request is passed to storage. */
export function TaskFeedback({
  result,
  defaultCategory = 'other',
  onSubmit,
  compact = false,
}: TaskFeedbackProps) {
  const [rating, setRating] = useState<FeedbackRating>();
  const [category, setCategory] = useState<FeedbackCategory>(defaultCategory);
  const [note, setNote] = useState('');
  const [submitted, setSubmitted] = useState(false);
  const [storageStatus, setStorageStatus] = useState<FeedbackStorageStatus>(() =>
    getFeedbackStorageStatus(),
  );
  const titleId = useId();
  const noteId = useId();
  const categoryId = useId();

  if (submitted) {
    const savedInBrowser = storageStatus.persistence === 'saved-in-browser';
    return (
      <section
        className={`task-feedback is-complete${compact ? ' is-compact' : ''}`}
        role="status"
        aria-live="polite"
        aria-atomic="true"
      >
        <strong>
          {savedInBrowser
            ? 'Thanks — feedback saved locally in this browser.'
            : 'Thanks — feedback kept for this session only.'}
        </strong>
        <p>
          {savedInBrowser
            ? 'Your rating, workflow area, and optional note were saved in this browser. Metadata is not automatically attached.'
            : 'Browser storage was unavailable, so this response remains in memory for this session only. Metadata is not automatically attached.'}
        </p>
        <p>
          The optional note is user-entered and may contain information. Keep it general and do not
          include confidential material in notes.
        </p>
      </section>
    );
  }

  const submit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (!rating) return;
    onSubmit({ result, rating, category, note });
    setStorageStatus(getFeedbackStorageStatus());
    setSubmitted(true);
  };

  return (
    <section
      className={`task-feedback${compact ? ' is-compact' : ''}`}
      aria-labelledby={titleId}
      data-testid="task-feedback"
    >
      <div className="task-feedback-heading">
        <div>
          <span className="studio-eyebrow">DESIGN-PARTNER SIGNAL</span>
          <h3 id={titleId}>Help improve this workflow</h3>
        </div>
        <span className="task-feedback-local-badge">Local only</span>
      </div>
      <p className="task-feedback-intro">
        {result === 'blocked'
          ? 'Was this boundary useful for the workbook task you had in mind?'
          : 'A quick signal helps prioritize the next workbook workflows to improve.'}
      </p>
      <form onSubmit={submit}>
        <fieldset className="task-feedback-rating-fieldset">
          <legend>{result === 'blocked' ? 'How did the block feel?' : 'How did it go?'}</legend>
          <div className="task-feedback-rating-list" role="group" aria-label="Outcome rating">
            {ratingLabels[result].map((option) => (
              <button
                type="button"
                key={option.value}
                className="task-feedback-rating"
                aria-pressed={rating === option.value}
                aria-label={option.label}
                onClick={() => setRating(option.value)}
              >
                {option.label}
              </button>
            ))}
          </div>
        </fieldset>
        <div className="task-feedback-fields">
          <div className="task-feedback-field">
            <label htmlFor={categoryId}>Workflow area</label>
            <select
              id={categoryId}
              value={category}
              onChange={(event) => setCategory(event.target.value as FeedbackCategory)}
            >
              {categoryOptions.map(([value, label]) => (
                <option value={value} key={value}>
                  {label}
                </option>
              ))}
            </select>
          </div>
          <div className="task-feedback-field">
            <label htmlFor={noteId}>Optional note</label>
            <textarea
              id={noteId}
              value={note}
              onChange={(event) => setNote(event.target.value)}
              maxLength={240}
              rows={2}
              aria-describedby={`${noteId}-privacy`}
              placeholder="Keep it general, such as ‘I needed a clearer preview.’"
            />
          </div>
        </div>
        <p className="task-feedback-privacy" id={`${noteId}-privacy`}>
          Metadata is not automatically attached. Optional notes are user-entered and may contain
          information; keep them generic and do not include workbook contents, file names, cell
          values, prompts, API keys, or other confidential material. Notes are limited to 240
          characters and common credential patterns are redacted before local storage.
        </p>
        <button type="submit" className="btn btn-secondary btn-sm" disabled={!rating}>
          Save local feedback
        </button>
      </form>
    </section>
  );
}
