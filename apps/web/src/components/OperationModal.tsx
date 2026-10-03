import React, { useState } from 'react';
import { type Workbook, indexToColumn } from '@excel-agent/engine';

interface OperationModalProps {
  isOpen: boolean;
  onClose: () => void;
  workbook: Workbook;
  activeSheetName: string;
  onExecute: (name: string, input: Record<string, unknown>) => void;
}

const OPERATIONS = [
  { id: 'format_dates', label: 'Format Dates (normalize date formats)' },
  { id: 'normalize_text', label: 'Normalize Text (trim, titlecase, uppercase, lowercase)' },
  { id: 'sort_range', label: 'Sort Range (ascending or descending)' },
  { id: 'filter_rows', label: 'Filter Rows (keep matching rows)' },
  { id: 'find_replace', label: 'Find & Replace' },
  { id: 'delete_duplicates', label: 'Delete Duplicate Rows' },
  { id: 'rename_column', label: 'Rename Column Header' },
  { id: 'delete_column', label: 'Delete Column' },
  { id: 'add_column', label: 'Add New Column' },
  { id: 'set_cells', label: 'Set Cell Value' },
];

type FilterOperator =
  | 'equals'
  | 'not_equals'
  | 'contains'
  | 'starts_with'
  | 'ends_with'
  | 'is_blank'
  | 'is_not_blank'
  | 'gt'
  | 'lt';

export const OperationModal: React.FC<OperationModalProps> = ({
  isOpen,
  onClose,
  workbook,
  activeSheetName,
  onExecute,
}) => {
  if (!isOpen) return null;

  const currentSheet =
    workbook.sheets.find((s) => s.name === activeSheetName) || workbook.sheets[0];
  const totalCols = currentSheet ? Math.max(...currentSheet.rows.map((r) => r.length), 0) : 0;
  const colLetters = Array.from({ length: totalCols }).map((_, i) => indexToColumn(i));

  const [selectedOp, setSelectedOp] = useState('format_dates');

  // Form states
  const [col, setCol] = useState(colLetters[0] || 'A');
  const [dateFormat, setDateFormat] = useState<'YYYY-MM-DD' | 'MM/DD/YYYY' | 'DD/MM/YYYY'>(
    'YYYY-MM-DD',
  );
  const [textTransform, setTextTransform] = useState<
    'trim' | 'lowercase' | 'uppercase' | 'titlecase'
  >('trim');
  const [sortDir, setSortDir] = useState<'asc' | 'desc'>('asc');
  const [filterCond, setFilterCond] = useState<FilterOperator>('equals');
  const [filterVal, setFilterVal] = useState('');
  const [findText, setFindText] = useState('');
  const [replaceText, setReplaceText] = useState('');
  const [newColName, setNewColName] = useState('');
  const [fillValue, setFillValue] = useState('');
  const [cellCoord, setCellCoord] = useState('A2');
  const [cellVal, setCellVal] = useState('');

  const handleRun = () => {
    if (!currentSheet) return;
    const sheet = currentSheet.name;

    switch (selectedOp) {
      case 'format_dates':
        onExecute('format_dates', {
          sheet,
          column: col,
          format: dateFormat,
          headerRow: 1,
        });
        break;

      case 'normalize_text':
        onExecute('normalize_text', {
          sheet,
          columns: [col],
          trim: true,
          collapseWhitespace: true,
          case:
            textTransform === 'titlecase'
              ? 'title'
              : textTransform === 'uppercase'
                ? 'upper'
                : textTransform === 'lowercase'
                  ? 'lower'
                  : 'none',
          headerRow: 1,
        });
        break;

      case 'sort_range':
        onExecute('sort_range', {
          sheet,
          column: col,
          direction: sortDir,
          startRow: 2,
          startColumn: 'A',
        });
        break;

      case 'filter_rows':
        onExecute('filter_rows', {
          sheet,
          column: col,
          operator: filterCond,
          value: filterVal,
          headerRow: 1,
        });
        break;

      case 'find_replace':
        onExecute('find_replace', {
          sheet,
          find: findText,
          replace: replaceText,
          matchCase: false,
          wholeCell: false,
          includeFormulas: false,
        });
        break;

      case 'delete_duplicates':
        onExecute('delete_duplicates', {
          sheet,
          columns: colLetters.length > 0 ? colLetters : ['A'],
          headerRow: 1,
          keep: 'first',
        });
        break;

      case 'rename_column':
        onExecute('rename_column', {
          sheet,
          column: col,
          newName: newColName || 'New Header',
          headerRow: 1,
        });
        break;

      case 'delete_column':
        onExecute('delete_column', {
          sheet,
          column: col,
        });
        break;

      case 'add_column': {
        const insertCol = colLetters[colLetters.length - 1] || 'A';
        onExecute('add_column', {
          sheet,
          column: insertCol,
          headerName: newColName || 'New Column',
          defaultValue: isNaN(Number(fillValue)) ? fillValue : Number(fillValue),
          headerRow: 1,
        });
        break;
      }

      case 'set_cells': {
        const match = cellCoord.toUpperCase().match(/^([A-Z]+)(\d+)$/);
        const targetCol = match ? match[1]! : 'A';
        const targetRow = match ? parseInt(match[2]!, 10) : 2;
        onExecute('set_cells', {
          sheet,
          cells: [
            {
              row: targetRow,
              column: targetCol,
              value: isNaN(Number(cellVal)) ? cellVal : Number(cellVal),
            },
          ],
        });
        break;
      }
    }

    onClose();
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-card" onClick={(e) => e.stopPropagation()}>
        <div className="modal-header">
          <div className="modal-title">Run Engine Operation</div>
          <button className="btn btn-ghost btn-sm" onClick={onClose}>
            ✕
          </button>
        </div>

        <div className="modal-body">
          <div className="form-group">
            <label className="form-label">Select Operation</label>
            <select
              className="select-input"
              value={selectedOp}
              onChange={(e) => setSelectedOp(e.target.value)}
            >
              {OPERATIONS.map((op) => (
                <option key={op.id} value={op.id}>
                  {op.label}
                </option>
              ))}
            </select>
          </div>

          {/* Operation Specific Inputs */}
          {selectedOp === 'format_dates' && (
            <>
              <div className="form-group">
                <label className="form-label">Target Column</label>
                <select
                  className="select-input"
                  value={col}
                  onChange={(e) => setCol(e.target.value)}
                >
                  {colLetters.map((c) => (
                    <option key={c} value={c}>
                      Column {c}
                    </option>
                  ))}
                </select>
              </div>

              <div className="form-group">
                <label className="form-label">Desired Date Format</label>
                <select
                  className="select-input"
                  value={dateFormat}
                  onChange={(e) =>
                    setDateFormat(e.target.value as 'YYYY-MM-DD' | 'MM/DD/YYYY' | 'DD/MM/YYYY')
                  }
                >
                  <option value="YYYY-MM-DD">ISO standard (YYYY-MM-DD)</option>
                  <option value="MM/DD/YYYY">US standard (MM/DD/YYYY)</option>
                  <option value="DD/MM/YYYY">European standard (DD/MM/YYYY)</option>
                </select>
              </div>
            </>
          )}

          {selectedOp === 'normalize_text' && (
            <>
              <div className="form-group">
                <label className="form-label">Target Column</label>
                <select
                  className="select-input"
                  value={col}
                  onChange={(e) => setCol(e.target.value)}
                >
                  {colLetters.map((c) => (
                    <option key={c} value={c}>
                      Column {c}
                    </option>
                  ))}
                </select>
              </div>

              <div className="form-group">
                <label className="form-label">Transform Type</label>
                <select
                  className="select-input"
                  value={textTransform}
                  onChange={(e) =>
                    setTextTransform(
                      e.target.value as 'trim' | 'lowercase' | 'uppercase' | 'titlecase',
                    )
                  }
                >
                  <option value="trim">Trim Whitespace (leading/trailing)</option>
                  <option value="titlecase">Title Case (Capitalize Every Word)</option>
                  <option value="uppercase">UPPERCASE</option>
                  <option value="lowercase">lowercase</option>
                </select>
              </div>
            </>
          )}

          {selectedOp === 'sort_range' && (
            <>
              <div className="form-group">
                <label className="form-label">Sort Column</label>
                <select
                  className="select-input"
                  value={col}
                  onChange={(e) => setCol(e.target.value)}
                >
                  {colLetters.map((c) => (
                    <option key={c} value={c}>
                      Column {c}
                    </option>
                  ))}
                </select>
              </div>

              <div className="form-group">
                <label className="form-label">Sort Direction</label>
                <select
                  className="select-input"
                  value={sortDir}
                  onChange={(e) => setSortDir(e.target.value as 'asc' | 'desc')}
                >
                  <option value="asc">Ascending (A to Z / Low to High)</option>
                  <option value="desc">Descending (Z to A / High to Low)</option>
                </select>
              </div>
            </>
          )}

          {selectedOp === 'filter_rows' && (
            <>
              <div className="form-group">
                <label className="form-label">Filter Column</label>
                <select
                  className="select-input"
                  value={col}
                  onChange={(e) => setCol(e.target.value)}
                >
                  {colLetters.map((c) => (
                    <option key={c} value={c}>
                      Column {c}
                    </option>
                  ))}
                </select>
              </div>

              <div className="form-group">
                <label className="form-label">Operator</label>
                <select
                  className="select-input"
                  value={filterCond}
                  onChange={(e) => setFilterCond(e.target.value as FilterOperator)}
                >
                  <option value="equals">Equals</option>
                  <option value="not_equals">Does Not Equal</option>
                  <option value="contains">Contains</option>
                  <option value="starts_with">Starts With</option>
                  <option value="ends_with">Ends With</option>
                  <option value="gt">Greater Than (&gt;)</option>
                  <option value="lt">Less Than (&lt;)</option>
                  <option value="is_blank">Is Blank</option>
                  <option value="is_not_blank">Is Not Blank</option>
                </select>
              </div>

              <div className="form-group">
                <label className="form-label">Comparison Value</label>
                <input
                  type="text"
                  className="form-input"
                  value={filterVal}
                  onChange={(e) => setFilterVal(e.target.value)}
                  placeholder="e.g. Completed or 1000"
                />
              </div>
            </>
          )}

          {selectedOp === 'find_replace' && (
            <>
              <div className="form-group">
                <label className="form-label">Find Text</label>
                <input
                  type="text"
                  className="form-input"
                  value={findText}
                  onChange={(e) => setFindText(e.target.value)}
                  placeholder="Text to find"
                />
              </div>
              <div className="form-group">
                <label className="form-label">Replace With</label>
                <input
                  type="text"
                  className="form-input"
                  value={replaceText}
                  onChange={(e) => setReplaceText(e.target.value)}
                  placeholder="Replacement text"
                />
              </div>
            </>
          )}

          {selectedOp === 'delete_duplicates' && (
            <p style={{ fontSize: '13px', color: 'var(--text-muted)' }}>
              This operation scans all rows in the active sheet ({currentSheet?.name}) across all
              columns ({colLetters.join(', ')}) and removes duplicate rows, keeping the first
              occurrence.
            </p>
          )}

          {selectedOp === 'rename_column' && (
            <>
              <div className="form-group">
                <label className="form-label">Select Column</label>
                <select
                  className="select-input"
                  value={col}
                  onChange={(e) => setCol(e.target.value)}
                >
                  {colLetters.map((c) => (
                    <option key={c} value={c}>
                      Column {c}
                    </option>
                  ))}
                </select>
              </div>
              <div className="form-group">
                <label className="form-label">New Header Name</label>
                <input
                  type="text"
                  className="form-input"
                  value={newColName}
                  onChange={(e) => setNewColName(e.target.value)}
                  placeholder="e.g. Total Revenue"
                />
              </div>
            </>
          )}

          {selectedOp === 'delete_column' && (
            <div className="form-group">
              <label className="form-label">Column to Delete</label>
              <select className="select-input" value={col} onChange={(e) => setCol(e.target.value)}>
                {colLetters.map((c) => (
                  <option key={c} value={c}>
                    Column {c}
                  </option>
                ))}
              </select>
            </div>
          )}

          {selectedOp === 'add_column' && (
            <>
              <div className="form-group">
                <label className="form-label">Column Name</label>
                <input
                  type="text"
                  className="form-input"
                  value={newColName}
                  onChange={(e) => setNewColName(e.target.value)}
                  placeholder="e.g. Tax Rate"
                />
              </div>
              <div className="form-group">
                <label className="form-label">Initial Default Value</label>
                <input
                  type="text"
                  className="form-input"
                  value={fillValue}
                  onChange={(e) => setFillValue(e.target.value)}
                  placeholder="e.g. 0.08 or N/A"
                />
              </div>
            </>
          )}

          {selectedOp === 'set_cells' && (
            <>
              <div className="form-group">
                <label className="form-label">Cell Address (e.g. C3)</label>
                <input
                  type="text"
                  className="form-input"
                  value={cellCoord}
                  onChange={(e) => setCellCoord(e.target.value)}
                />
              </div>
              <div className="form-group">
                <label className="form-label">New Cell Value</label>
                <input
                  type="text"
                  className="form-input"
                  value={cellVal}
                  onChange={(e) => setCellVal(e.target.value)}
                  placeholder="New value or number"
                />
              </div>
            </>
          )}
        </div>

        <div className="modal-footer">
          <button className="btn btn-ghost btn-sm" onClick={onClose}>
            Cancel
          </button>
          <button className="btn btn-primary btn-sm" onClick={handleRun}>
            Execute Operation
          </button>
        </div>
      </div>
    </div>
  );
};
