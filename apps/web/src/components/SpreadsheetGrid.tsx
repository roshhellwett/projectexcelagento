import React, { useState, useEffect } from 'react';
import { type Workbook, type Sheet, type Cell, indexToColumn } from '@excel-agent/engine';

interface SpreadsheetGridProps {
  workbook: Workbook;
  activeSheetName: string;
  onSelectSheet: (sheetName: string) => void;
  recentChangedCells: Set<string>; // formatted as `${sheetName}:${row}:${colLetter}`
  searchHighlightCells?: Set<string>;
  onQuickSort?: (columnLetter: string, direction: 'asc' | 'desc') => void;
  onSelectCell?: (coord: { row: number; column: string; value: unknown }) => void;
  onFileDrop?: (file: File) => void;
}

export const SpreadsheetGrid: React.FC<SpreadsheetGridProps> = ({
  workbook,
  activeSheetName,
  onSelectSheet,
  recentChangedCells,
  searchHighlightCells = new Set(),
  onQuickSort,
  onSelectCell,
  onFileDrop,
}) => {
  const currentSheet: Sheet = workbook.sheets.find((s) => s.name === activeSheetName) ||
    workbook.sheets[0] || { name: 'Sheet1', rows: [] };

  const [selectedCell, setSelectedCell] = useState<{
    row: number; // 1-indexed
    colIdx: number; // 0-indexed
  }>({ row: 1, colIdx: 0 });

  const [isDragOver, setIsDragOver] = useState(false);

  const totalCols = Math.max(...currentSheet.rows.map((r) => r.length), 0);

  const selectedColLetter = indexToColumn(selectedCell.colIdx);
  const selectedCellCoord = `${selectedColLetter}${selectedCell.row}`;

  const activeRowCells = currentSheet.rows[selectedCell.row - 1] || [];
  const currentCell: Cell | undefined = activeRowCells[selectedCell.colIdx];

  const cellValue = currentCell?.value;
  const cellFormula = currentCell?.formula;

  useEffect(() => {
    if (onSelectCell && currentCell) {
      onSelectCell({
        row: selectedCell.row,
        column: selectedColLetter,
        value: currentCell.value,
      });
    }
  }, [selectedCell, currentCell, onSelectCell, selectedColLetter]);

  let cellTypeStr = 'empty';
  if (cellFormula) cellTypeStr = 'formula';
  else if (typeof cellValue === 'number') cellTypeStr = 'number';
  else if (typeof cellValue === 'boolean') cellTypeStr = 'boolean';
  else if (typeof cellValue === 'string') {
    if (/^\d{4}-\d{2}-\d{2}$/.test(cellValue) || /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(cellValue)) {
      cellTypeStr = 'date';
    } else {
      cellTypeStr = 'string';
    }
  }

  // Drag and drop handlers
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
    const file = e.dataTransfer.files?.[0];
    if (file && onFileDrop) {
      onFileDrop(file);
    }
  };

  return (
    <div
      className="grid-panel"
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
    >
      {/* Drag & Drop Visual Overlay */}
      {isDragOver && (
        <div className="grid-drag-overlay">
          <div className="drag-overlay-card">
            <div className="drag-icon-bounce">
              <svg
                width="24"
                height="24"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <rect x="3" y="3" width="18" height="18" rx="2" ry="2" />
                <line x1="3" y1="9" x2="21" y2="9" />
                <line x1="3" y1="15" x2="21" y2="15" />
                <line x1="9" y1="3" x2="9" y2="21" />
                <line x1="15" y1="3" x2="15" y2="21" />
              </svg>
            </div>
            <h3 style={{ fontSize: '15px', fontWeight: 600 }}>Drop your Excel or CSV file here</h3>
            <p style={{ fontSize: '12px', color: 'var(--text-muted)' }}>
              Excel Agent will inspect and parse your spreadsheet locally
            </p>
          </div>
        </div>
      )}

      {/* Formula Bar */}
      <div className="formula-bar">
        <div className="cell-coord-badge">{selectedCellCoord}</div>
        <div className="formula-type-badge">{cellTypeStr}</div>
        <div className="formula-input-display">
          {cellFormula ? (
            <span className="cell-val-formula">={cellFormula}</span>
          ) : cellValue !== null && cellValue !== undefined ? (
            String(cellValue)
          ) : (
            <span style={{ color: 'var(--text-subtle)' }}>(empty)</span>
          )}
        </div>
      </div>

      {/* Grid Scroll Table */}
      <div className="grid-scroll-wrapper">
        <table className="spreadsheet-table">
          <thead>
            <tr>
              <th className="corner-cell" />
              {Array.from({ length: totalCols }).map((_, cIdx) => {
                const colLetter = indexToColumn(cIdx);

                return (
                  <th key={colLetter} className="column-header">
                    <div className="col-header-inner">
                      <span className="col-letter">{colLetter}</span>

                      {onQuickSort && (
                        <div style={{ display: 'flex', gap: '2px' }}>
                          <button
                            className="btn btn-ghost btn-sm"
                            style={{ padding: '0 3px', height: '18px', fontSize: '10px' }}
                            title={`Sort Column ${colLetter} Ascending`}
                            onClick={(e) => {
                              e.stopPropagation();
                              onQuickSort(colLetter, 'asc');
                            }}
                          >
                            ▲
                          </button>
                          <button
                            className="btn btn-ghost btn-sm"
                            style={{ padding: '0 3px', height: '18px', fontSize: '10px' }}
                            title={`Sort Column ${colLetter} Descending`}
                            onClick={(e) => {
                              e.stopPropagation();
                              onQuickSort(colLetter, 'desc');
                            }}
                          >
                            ▼
                          </button>
                        </div>
                      )}
                    </div>
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {currentSheet.rows.map((rowCells, rIdx) => {
              const rowNumber = rIdx + 1;
              const isHeaderRow = rowNumber === 1;

              return (
                <tr key={rowNumber}>
                  <td className="row-index-cell">{rowNumber}</td>
                  {Array.from({ length: totalCols }).map((_, cIdx) => {
                    const colLetter = indexToColumn(cIdx);
                    const cell = rowCells[cIdx];
                    const isSelected =
                      selectedCell.row === rowNumber && selectedCell.colIdx === cIdx;
                    const cellKey = `${currentSheet.name}:${rowNumber}:${colLetter}`;
                    const isDiffChanged = recentChangedCells.has(cellKey);
                    const isSearchMatch = searchHighlightCells.has(cellKey);

                    const val = cell?.value;
                    let displayVal: React.ReactNode = '';
                    let cellClass = 'data-cell';

                    if (isHeaderRow) cellClass += ' cell-header-row';
                    if (isSelected) cellClass += ' selected';
                    if (isDiffChanged) cellClass += ' diff-changed';
                    if (isSearchMatch) cellClass += ' cell-search-match';

                    if (cell?.formula) {
                      displayVal = <span className="cell-val-formula">={cell.formula}</span>;
                    } else if (typeof val === 'number') {
                      cellClass += ' cell-val-number';
                      displayVal = val.toLocaleString();
                    } else if (val === null || val === undefined || val === '') {
                      cellClass += ' cell-val-null';
                      displayVal = '';
                    } else {
                      displayVal = String(val);
                    }

                    return (
                      <td
                        key={colLetter}
                        className={cellClass}
                        onClick={() => setSelectedCell({ row: rowNumber, colIdx: cIdx })}
                        title={`${colLetter}${rowNumber}: ${val ?? '(empty)'}`}
                      >
                        {displayVal}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* Multi-Sheet Tabs */}
      <div className="sheet-tabs-bar">
        {workbook.sheets.map((sheet) => {
          const isActive = sheet.name === currentSheet.name;
          const count = sheet.rows.length;

          return (
            <button
              key={sheet.name}
              className={`sheet-tab ${isActive ? 'active' : ''}`}
              onClick={() => onSelectSheet(sheet.name)}
            >
              <span>{sheet.name}</span>
              <span className="sheet-tab-count">{count} rows</span>
            </button>
          );
        })}
      </div>
    </div>
  );
};
