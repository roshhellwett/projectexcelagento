import { useMemo, useRef } from 'react';
import { FileSpreadsheet, Upload } from 'lucide-react';
import type { CompactColumnProfile } from '@excel-agent/agent';
import { workspaceWorkflows } from '../lib/workflows.js';

interface WorkbookQuickStartProps {
  fileName: string;
  sheetName: string;
  isSample: boolean;
  profiles: CompactColumnProfile[];
  isProcessing: boolean;
  onUpload: (file: File) => void;
  onUseSample: () => void;
  onDraft: (prompt: string) => void;
}

/** A first task uses the same composer, preview, and confirmation path as every other request. */
export function WorkbookQuickStart({
  fileName,
  sheetName,
  isSample,
  profiles,
  isProcessing,
  onUpload,
  onUseSample,
  onDraft,
}: WorkbookQuickStartProps) {
  const uploadInput = useRef<HTMLInputElement>(null);
  const starters = useMemo(
    () =>
      workspaceWorkflows(profiles).filter((workflow) =>
        ['duplicates', 'overview', 'missing'].includes(workflow.id),
      ),
    [profiles],
  );
  const hasData = profiles.some((column) => column.nonBlankCount > 0);

  return (
    <section className="workbook-quick-start" aria-labelledby="workbook-start-title">
      <div className="workbook-start-heading">
        <div>
          <h2 id="workbook-start-title">Get your first workbook result</h2>
          <p>
            <strong>{fileName}</strong> · {sheetName}
            <br />
            {hasData
              ? isSample
                ? 'Practice with the sample, or upload your own Excel or CSV file.'
                : 'Your workbook is ready. Choose a prompt for this sheet, or write your own.'
              : 'This sheet has no data rows. Choose another sheet, upload a file, or try the sample.'}
          </p>
        </div>
        <div className="workbook-start-actions">
          <button
            type="button"
            className="btn btn-secondary btn-sm"
            disabled={isProcessing}
            onClick={() => uploadInput.current?.click()}
          >
            <Upload size={14} /> Upload your workbook
          </button>
          {!isSample && (
            <button
              type="button"
              className="btn btn-ghost btn-sm"
              disabled={isProcessing}
              onClick={onUseSample}
            >
              <FileSpreadsheet size={14} /> Use sample workbook
            </button>
          )}
          <input
            ref={uploadInput}
            type="file"
            hidden
            accept=".xlsx,.xls,.csv"
            aria-label="Upload a workbook to get started"
            onChange={(event) => {
              const file = event.target.files?.[0];
              if (file) onUpload(file);
              event.target.value = '';
            }}
          />
        </div>
      </div>
      {hasData && (
        <>
          <div className="workbook-start-prompts" role="group" aria-label="First task prompts">
            {starters.map((workflow) => (
              <button
                type="button"
                key={workflow.id}
                className={`btn btn-sm ${workflow.id === 'duplicates' ? 'btn-primary' : 'btn-secondary'}`}
                disabled={isProcessing}
                onClick={() => onDraft(workflow.prompt)}
                title={`Draft: ${workflow.prompt}`}
              >
                {isSample && workflow.id === 'duplicates' ? 'Try sample cleanup' : workflow.title}
              </button>
            ))}
          </div>
          <p className="workbook-start-help">
            Choose a prompt, edit or send it, then review the preview before applying. These
            starters work without an AI key.
          </p>
        </>
      )}
    </section>
  );
}
