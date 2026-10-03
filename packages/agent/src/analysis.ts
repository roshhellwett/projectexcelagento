import { type Workbook, type Sheet, type CellValue, indexToColumn } from '@excel-agent/engine';

export interface ProposedAction {
  name: string;
  args: Record<string, unknown>;
  explanation: string;
  category: 'format' | 'transform' | 'filter' | 'columns' | 'structure';
}

export interface SheetAudit {
  sheetName: string;
  totalRows: number;
  totalCols: number;
  headers: string[];
  findings: string[];
  suggestions: { prompt: string; action: ProposedAction }[];
}

export interface CellMatch {
  sheet: string;
  row: number;
  column: string;
  value: CellValue;
  text: string;
}

export interface ColumnMetadata {
  letter: string;
  index: number;
  rawName: string;
  cleanName: string;
  isNumeric: boolean;
  isDate: boolean;
  nonBlankCount: number;
  numericValues: number[];
  sum?: number;
  avg?: number;
  min?: number;
  max?: number;
  distinct: Map<string, number>;
}

export function isFuzzyMatch(word: string, target: string): boolean {
  if (word.length < 3 || target.length < 3) return word === target;
  if (Math.abs(word.length - target.length) > 2) return false;
  if (word.split('').sort().join('') === target.split('').sort().join('')) return true;

  const m = word.length;
  const n = target.length;
  const d: number[][] = Array.from({ length: m + 1 }, () => Array(n + 1).fill(0));
  for (let a = 0; a <= m; a++) d[a]![0] = a;
  for (let b = 0; b <= n; b++) d[0]![b] = b;
  for (let a = 1; a <= m; a++) {
    for (let b = 1; b <= n; b++) {
      d[a]![b] =
        word[a - 1] === target[b - 1]
          ? d[a - 1]![b - 1]!
          : Math.min(d[a - 1]![b]! + 1, d[a]![b - 1]! + 1, d[a - 1]![b - 1]! + 1);
    }
  }
  return d[m]![n]! <= (target.length > 5 ? 2 : 1);
}

export function searchCellsInSheet(sheet: Sheet, query: string): CellMatch[] {
  const q = query.trim().toLowerCase();
  if (!q) return [];

  const matches: CellMatch[] = [];
  const fuzzyMatches: CellMatch[] = [];

  sheet.rows.forEach((row, rIdx) => {
    const rowNum = rIdx + 1;
    row.forEach((cell, cIdx) => {
      const colLetter = indexToColumn(cIdx);
      const val = cell?.value;
      const strVal = val !== null && val !== undefined ? String(val) : '';
      const lower = strVal.toLowerCase();

      if (lower.includes(q)) {
        matches.push({
          sheet: sheet.name,
          row: rowNum,
          column: colLetter,
          value: val,
          text: strVal,
        });
      } else {
        const words = lower.split(/[\s,._-]+/);
        if (words.some((w) => isFuzzyMatch(w, q))) {
          fuzzyMatches.push({
            sheet: sheet.name,
            row: rowNum,
            column: colLetter,
            value: val,
            text: strVal,
          });
        }
      }
    });
  });

  return matches.length > 0 ? matches : fuzzyMatches;
}

export function getColumnProfiles(sheet: Sheet): ColumnMetadata[] {
  const totalCols = Math.max(...sheet.rows.map((r) => r.length), 0);
  const headerRow = sheet.rows[0] || [];
  const dataRows = sheet.rows.slice(1);

  const columns: ColumnMetadata[] = [];

  for (let c = 0; c < totalCols; c++) {
    const letter = indexToColumn(c);
    const rawVal = headerRow[c]?.value;
    const rawName =
      rawVal !== null && rawVal !== undefined ? String(rawVal).trim() : `Column ${letter}`;
    const cleanName = rawName.toLowerCase();

    const distinct = new Map<string, number>();
    const numericValues: number[] = [];
    let dateCount = 0;
    let nonBlankCount = 0;

    for (const r of dataRows) {
      const val = r[c]?.value;
      if (val !== null && val !== undefined && String(val).trim() !== '') {
        nonBlankCount++;
        const strVal = String(val).trim();
        distinct.set(strVal, (distinct.get(strVal) || 0) + 1);

        const cleanNumStr = strVal.replace(/,/g, '');
        const num = typeof val === 'number' ? val : Number(cleanNumStr);
        if (!isNaN(num) && cleanNumStr !== '') {
          numericValues.push(num);
        }

        if (
          /^\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(strVal) ||
          /^\d{1,2}[-/]\d{1,2}[-/]\d{4}/.test(strVal)
        ) {
          dateCount++;
        }
      }
    }

    const isNumeric = numericValues.length > 0 && numericValues.length >= nonBlankCount * 0.7;
    const isDate = dateCount > 0 && dateCount >= nonBlankCount * 0.6;

    let sum = undefined;
    let avg = undefined;
    let min = undefined;
    let max = undefined;

    if (isNumeric && numericValues.length > 0) {
      sum = numericValues.reduce((a, b) => a + b, 0);
      avg = sum / numericValues.length;
      min = Math.min(...numericValues);
      max = Math.max(...numericValues);
    }

    columns.push({
      letter,
      index: c,
      rawName,
      cleanName,
      isNumeric,
      isDate,
      nonBlankCount,
      numericValues,
      sum,
      avg,
      min,
      max,
      distinct,
    });
  }

  return columns;
}

const CONCEPT_SYNONYMS: Record<string, string[]> = {
  stock: ['stock', 'available', 'inventory', 'balance', 'godown', 'store', 'stocks', 'curr_stock'],
  quantity: [
    'quantity',
    'qty',
    'units',
    'items',
    'count',
    'pieces',
    'pcs',
    'volume',
    'kitna',
    'amount',
  ],
  date: ['date', 'time', 'timestamp', 'created', 'updated', 'day', 'din', 'tarikh', 'when'],
  status: ['status', 'state', 'condition', 'action', 'type', 'stage'],
  user: ['user', 'handled', 'name', 'person', 'agent', 'by', 'employee', 'who', 'staff', 'owner'],
  customer: ['customer', 'client', 'buyer', 'account', 'company', 'organization'],
  price: [
    'price',
    'cost',
    'amount',
    'usd',
    'inr',
    'rate',
    'revenue',
    'sales',
    'value',
    'total',
    'target',
  ],
  group: ['group', 'micron', 'mic', 'specification', 'width', 'size', 'color', 'shade', 'category'],
};

export function resolveColumn(query: string, columns: ColumnMetadata[]): ColumnMetadata | null {
  const q = query.trim().toLowerCase();
  if (!q) return null;

  // 0. Bare column reference (e.g. "A", "AB")
  if (/^[a-z]{1,2}$/.test(q)) {
    const direct = columns.find((c) => c.letter.toLowerCase() === q);
    if (direct) return direct;
  }

  // 1. Direct match by letter (e.g. "column A", "col B", "in C")
  const letterMatch =
    q.match(/\b(?:column|col)\s+([a-z]{1,2})\b/i) || q.match(/\bin\s+([a-z]{1,2})\b/i);
  if (letterMatch && letterMatch[1]) {
    const found = columns.find((c) => c.letter.toLowerCase() === letterMatch[1]?.toLowerCase());
    if (found) return found;
  }

  // 2. Exact or substring match in header name (only for meaningful query lengths)
  if (q.length >= 3) {
    for (const col of columns) {
      if (q === col.cleanName || q.includes(col.cleanName) || col.cleanName.includes(q)) {
        return col;
      }
    }
  }

  // 3. Token containment
  const queryTokens = q.split(/[\s,._/?!+-]+/).filter((t) => t.length > 2);
  for (const token of queryTokens) {
    for (const col of columns) {
      if (col.cleanName.includes(token)) {
        return col;
      }
    }
  }

  // 4. Synonym match
  for (const token of queryTokens) {
    for (const [concept, syns] of Object.entries(CONCEPT_SYNONYMS)) {
      if (concept === token || syns.includes(token)) {
        for (const col of columns) {
          if (col.cleanName.includes(concept) || syns.some((s) => col.cleanName.includes(s))) {
            return col;
          }
        }
      }
    }
  }

  // 5. Fuzzy match against header tokens
  for (const token of queryTokens) {
    for (const col of columns) {
      const headerTokens = col.cleanName.split(/[\s,._/()-]+/).filter((t) => t.length > 2);
      if (headerTokens.some((ht) => isFuzzyMatch(token, ht))) {
        return col;
      }
    }
  }

  return null;
}

export function auditSheet(sheet: Sheet): SheetAudit {
  const totalRows = sheet.rows.length;
  const totalCols = Math.max(...sheet.rows.map((r) => r.length), 0);
  const headerRow = sheet.rows[0] || [];
  const headers = headerRow.map((c, i) =>
    c?.value !== null && c?.value !== undefined ? String(c.value) : `Col ${indexToColumn(i)}`,
  );
  const allColumns = Array.from({ length: totalCols }, (_, i) => indexToColumn(i));

  const findings: string[] = [];
  const suggestions: { prompt: string; action: ProposedAction }[] = [];

  if (totalRows <= 1) {
    return {
      sheetName: sheet.name,
      totalRows,
      totalCols,
      headers,
      findings: ['Sheet has no data rows.'],
      suggestions: [],
    };
  }

  // Check for duplicate rows
  const seenRows = new Set<string>();
  let duplicateCount = 0;
  for (let r = 1; r < sheet.rows.length; r += 1) {
    const key = (sheet.rows[r] || []).map((c) => String(c?.value ?? '')).join('|~|');
    if (seenRows.has(key)) {
      duplicateCount += 1;
    } else {
      seenRows.add(key);
    }
  }

  if (duplicateCount > 0 && allColumns.length > 0) {
    findings.push(`Found ${duplicateCount} duplicate row(s).`);
    suggestions.push({
      prompt: 'Remove duplicate rows',
      action: {
        name: 'delete_duplicates',
        args: {
          sheet: sheet.name,
          columns: allColumns,
          headerRow: 1,
          keep: 'first',
        },
        explanation: `Delete ${duplicateCount} duplicate row(s), checking across all columns and keeping the first occurrence.`,
        category: 'structure',
      },
    });
  }

  // Scan columns for date or text irregularities
  for (let colIdx = 0; colIdx < totalCols; colIdx += 1) {
    const colLetter = indexToColumn(colIdx);
    const colHeader = headers[colIdx] || `Col ${colLetter}`;
    const colValues = sheet.rows.slice(1).map((r) => r[colIdx]?.value);

    // Date heuristic
    let dateLikeCount = 0;
    const dateFormatsSeen = new Set<string>();
    for (const v of colValues) {
      if (typeof v === 'string') {
        if (/^\d{4}-\d{2}-\d{2}$/.test(v)) {
          dateLikeCount += 1;
          dateFormatsSeen.add('ISO');
        } else if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(v)) {
          dateLikeCount += 1;
          dateFormatsSeen.add('Slash');
        } else if (/^\d{4}\/\d{1,2}\/\d{1,2}$/.test(v)) {
          dateLikeCount += 1;
          dateFormatsSeen.add('Slash-YMD');
        } else if (/^\d{1,2}-\d{1,2}-\d{4}$/.test(v)) {
          dateLikeCount += 1;
          dateFormatsSeen.add('Dash-DMY');
        }
      }
    }

    if (dateLikeCount >= 2) {
      if (dateFormatsSeen.size > 1) {
        findings.push(
          `Column "${colHeader}" (${colLetter}) has mixed date formats (${Array.from(dateFormatsSeen).join(', ')}).`,
        );
      }
      suggestions.push({
        prompt: `Normalize dates in ${colHeader} to ISO (YYYY-MM-DD)`,
        action: {
          name: 'format_dates',
          args: {
            sheet: sheet.name,
            column: colLetter,
            format: 'YYYY-MM-DD',
            headerRow: 1,
          },
          explanation: `Convert all dates in column ${colLetter} ("${colHeader}") into standard ISO YYYY-MM-DD format.`,
          category: 'format',
        },
      });
    }

    // Text whitespace & casing heuristic
    let untrimmedCount = 0;
    let mixedCaseCount = 0;
    for (const v of colValues) {
      if (typeof v === 'string' && v.trim().length > 0) {
        if (v.trim() !== v) untrimmedCount += 1;
        if (v === v.toLowerCase() || v === v.toUpperCase()) mixedCaseCount += 1;
      }
    }

    if (untrimmedCount > 0) {
      findings.push(
        `Column "${colHeader}" (${colLetter}) has ${untrimmedCount} untrimmed value(s).`,
      );
      suggestions.push({
        prompt: `Trim whitespace in ${colHeader}`,
        action: {
          name: 'normalize_text',
          args: {
            sheet: sheet.name,
            columns: [colLetter],
            trim: true,
            collapseWhitespace: true,
            case: 'none',
            headerRow: 1,
          },
          explanation: `Trim leading and trailing whitespace in column ${colLetter} ("${colHeader}").`,
          category: 'transform',
        },
      });
    } else if (mixedCaseCount >= 2 && dateLikeCount === 0) {
      suggestions.push({
        prompt: `Capitalize names in ${colHeader} to Title Case`,
        action: {
          name: 'normalize_text',
          args: {
            sheet: sheet.name,
            columns: [colLetter],
            trim: true,
            collapseWhitespace: true,
            case: 'title',
            headerRow: 1,
          },
          explanation: `Normalize text in column ${colLetter} ("${colHeader}") to proper Title Case.`,
          category: 'transform',
        },
      });
    }
  }

  // Suggest sorting by first column
  if (totalCols > 0) {
    const colLetter = indexToColumn(0);
    const colHeader = headers[0] || 'Column A';
    suggestions.push({
      prompt: `Sort rows by ${colHeader} (A-Z)`,
      action: {
        name: 'sort_range',
        args: {
          sheet: sheet.name,
          column: colLetter,
          direction: 'asc',
          startRow: 2,
          startColumn: 'A',
        },
        explanation: `Sort data rows by column ${colLetter} in ascending order.`,
        category: 'transform',
      },
    });
  }

  return {
    sheetName: sheet.name,
    totalRows,
    totalCols,
    headers,
    findings,
    suggestions: suggestions.slice(0, 6),
  };
}

/**
 * Autonomous In-Memory Spreadsheet Data Scientist Engine
 * Handles 90%+ of Excel workflows: filtering, row lookup, calculations, aggregations,
 * group breakdowns, sorting, cleaning, find-replace, column operations, and audits.
 */
export function analyzeSpreadsheetIntentAndData(
  userQuery: string,
  workbook: Workbook,
  activeSheetName: string,
): { message: string; proposedAction?: ProposedAction } {
  const currentSheet =
    workbook.sheets.find((s) => s.name === activeSheetName) || workbook.sheets[0];
  if (!currentSheet || currentSheet.rows.length === 0) {
    return {
      message: 'The current sheet is empty. Please upload or load an Excel spreadsheet to begin.',
    };
  }

  const q = userQuery.trim().toLowerCase();
  // Structural patterns (delete/rename/find-replace) are matched against the
  // original text so user-supplied names and replacement values keep their
  // casing; trigger checks below use the normalized copy.
  const raw = userQuery.trim();
  const columns = getColumnProfiles(currentSheet);
  const totalRows = currentSheet.rows.length;
  const dataRowsCount = Math.max(0, totalRows - 1);
  const allColumns = columns.map((c) => c.letter);

  // 0. EXPLICIT STRUCTURAL COMMANDS (highest precedence)
  // An unambiguous "rename column X to Y", "delete column X", or "replace X with
  // Y" must never be hijacked by a keyword heuristic below. Without this, a
  // column named "Order Date" turns "rename column C to Order Date" into a
  // date-format mutation, and "delete column Sort Priority" into a sort.
  const replaceMatch = raw.match(
    /(?:find\s+and\s+replace|find|replace|change)\s*['"]?([^'"]+?)['"]?\s*(?:and\s+)?(?:replace|replace\s+with|with|to)\s*(?:with\s+)?['"]?([^'"]+?)['"]?(?:\s+in\b.*)?$/i,
  );

  if (replaceMatch && replaceMatch[1] && replaceMatch[2]) {
    const find = replaceMatch[1].trim();
    const replace = replaceMatch[2].trim();
    return {
      message: `I've prepared a **Find & Replace** operation:\n\n• Replace: \`${find}\`\n• With: \`${replace}\`\n• Scope: Entire active sheet (${currentSheet.name})\n\nClick **Apply Changes** to execute this across all cells.`,
      proposedAction: {
        name: 'find_replace',
        args: {
          sheet: currentSheet.name,
          find,
          replace,
          matchCase: false,
          wholeCell: false,
          includeFormulas: false,
        },
        explanation: `Replace occurrences of "${find}" with "${replace}" in ${currentSheet.name}.`,
        category: 'transform',
      },
    };
  }

  const deleteColMatch = raw.match(
    /(?:delete|remove|drop)\s+(?:the\s+)?(?:column\s+)?([a-z0-9_\s/()]+?)(?:\s+column)?\s*$/i,
  );
  if (deleteColMatch && deleteColMatch[1]) {
    const targetCol = resolveColumn(deleteColMatch[1], columns);
    if (targetCol) {
      return {
        message: `I've prepared to remove column **${targetCol.rawName}** (${targetCol.letter}).\n\nReview the preview card and click **Apply Changes** to delete it.`,
        proposedAction: {
          name: 'delete_column',
          args: { sheet: currentSheet.name, column: targetCol.letter },
          explanation: `Delete column ${targetCol.letter} ("${targetCol.rawName}") from the sheet.`,
          category: 'columns',
        },
      };
    }
  }

  const renameColMatch = raw.match(
    /(?:rename)\s+(?:the\s+)?(?:column\s+)?([a-z0-9_\s/()]+?)\s+(?:to|as)\s+(.+)/i,
  );
  if (renameColMatch && renameColMatch[1] && renameColMatch[2]) {
    const targetCol = resolveColumn(renameColMatch[1], columns);
    const newName = renameColMatch[2].trim();
    if (targetCol) {
      return {
        message: `I've prepared to rename column **${targetCol.rawName}** (${targetCol.letter}) to **"${newName}"**.\n\nClick **Apply Changes** to update the header.`,
        proposedAction: {
          name: 'rename_column',
          args: { sheet: currentSheet.name, column: targetCol.letter, newName, headerRow: 1 },
          explanation: `Rename column ${targetCol.letter} to "${newName}".`,
          category: 'columns',
        },
      };
    }
  }

  // 0b. COLUMN STRUCTURE: add / fill / split / merge
  const addColMatch =
    raw.match(
      /(?:add|insert|create)\s+(?:an?\s+)?(?:new\s+)?(?:empty\s+)?column\s+(?:called\s+|named\s+)?['"]?([a-z0-9_\s-]+)['"]?/i,
    ) || raw.match(/(?:add|insert|create)\s+(?:a\s+)?(?:new\s+)?(?:empty\s+)?column/i);
  if (addColMatch) {
    const headerName = (addColMatch[1] || 'New Column').trim();
    const lastCol = columns[columns.length - 1]?.letter;
    const afterIndex = lastCol ? lastCol.charCodeAt(0) - 64 : 0;
    const insertLetter = indexToColumn(Math.max(afterIndex, 0));
    return {
      message: `I've prepared to add a new column **"${headerName}"** after the last existing column.\n\nClick **Apply Changes** to insert it.`,
      proposedAction: {
        name: 'add_column',
        args: {
          sheet: currentSheet.name,
          column: insertLetter,
          headerName,
          defaultValue: '',
          headerRow: 1,
        },
        explanation: `Add a new column "${headerName}" as column ${insertLetter}.`,
        category: 'columns',
      },
    };
  }

  const fillMatch = raw.match(
    /(?:fill|complete)\s+(?:the\s+)?(?:blank|empty|missing)(?:\s+(?:cells?|values?))?(?:\s+in\s+column\s+([a-z0-9_\s/()]+))?/i,
  );
  if (fillMatch) {
    const targetCol =
      (fillMatch[1] ? resolveColumn(fillMatch[1], columns) : null) ||
      columns.find((c) => c.nonBlankCount < currentSheet.rows.length - 1) ||
      columns[0]!;
    return {
      message: `I've prepared to fill blank cells in **${targetCol.rawName}** (${targetCol.letter}).\n\nClick **Apply Changes** to propagate the previous value forward.`,
      proposedAction: {
        name: 'fill_blanks',
        args: {
          sheet: currentSheet.name,
          column: targetCol.letter,
          strategy: 'forward',
          headerRow: 1,
        },
        explanation: `Fill blank cells in column ${targetCol.letter} using forward fill.`,
        category: 'transform',
      },
    };
  }

  const splitMatch = raw.match(
    /split\s+column\s+([a-z0-9_\s/()]+)\s+by\s+(['"]?[^'"]+?['"]?)\s*$/i,
  );
  if (splitMatch) {
    const targetCol = resolveColumn(splitMatch[1] ?? '', columns);
    const delimiter = splitMatch[2]!.replace(/^['"]|['"]$/g, '');
    if (targetCol && delimiter) {
      return {
        message: `I've prepared to split **${targetCol.rawName}** (${targetCol.letter}) on "${delimiter}".\n\nClick **Apply Changes** to expand it into new columns.`,
        proposedAction: {
          name: 'split_column',
          args: { sheet: currentSheet.name, column: targetCol.letter, delimiter, headerRow: 1 },
          explanation: `Split column ${targetCol.letter} by "${delimiter}".`,
          category: 'transform',
        },
      };
    }
  }

  const mergeMatch = raw.match(
    /merge\s+columns?\s+([a-z0-9_\s/(),&]+?)\s+(?:into|as|to)\s+['"]?([a-z0-9_\s-]+)['"]?/i,
  );
  if (mergeMatch) {
    const headerName = mergeMatch[2]!.trim();
    const parts = mergeMatch[1]!
      .split(/&|,|and/i)
      .map((p) => p.trim())
      .filter(Boolean);
    const resolved = parts.map((p) => resolveColumn(p, columns)).filter((c) => c !== null);
    if (resolved.length >= 2) {
      const letters = resolved.map((c) => c!.letter);
      return {
        message: `I've prepared to merge columns **${letters.join(', ')}** into **"${headerName}"**.\n\nClick **Apply Changes** to combine them.`,
        proposedAction: {
          name: 'merge_columns',
          args: {
            sheet: currentSheet.name,
            columns: letters,
            separator: ' ',
            headerName,
            headerRow: 1,
          },
          explanation: `Merge columns ${letters.join(', ')} into "${headerName}".`,
          category: 'transform',
        },
      };
    }
  }

  // 1. DUPLICATE REMOVAL
  if (
    q.includes('duplicate') ||
    q.includes('dedup') ||
    q.includes('unique') ||
    q.includes('dublicate') ||
    q.includes('repeated') ||
    q.includes('doublon')
  ) {
    const seen = new Set<string>();
    let dupCount = 0;
    for (let r = 1; r < currentSheet.rows.length; r++) {
      const key = (currentSheet.rows[r] || []).map((c) => String(c?.value ?? '')).join('|~|');
      if (seen.has(key)) dupCount++;
      else seen.add(key);
    }
    return {
      message: `I audited all **${dataRowsCount} data rows** in **${currentSheet.name}**.\n\n• Found **${dupCount} duplicate row(s)** across the dataset.\n\nI've generated a clean deduplication action for you. Review the card below and click **Apply Changes** to remove them instantly!`,
      proposedAction: {
        name: 'delete_duplicates',
        args: {
          sheet: currentSheet.name,
          columns: allColumns,
          headerRow: 1,
          keep: 'first',
        },
        explanation: `Remove ${dupCount} duplicate rows across all ${allColumns.length} columns, keeping the first occurrence.`,
        category: 'structure',
      },
    };
  }

  // 2. TEXT NORMALIZATION (TRIM / CASING)
  if (
    q.includes('trim') ||
    q.includes('whitespace') ||
    q.includes('clean') ||
    q.includes('space') ||
    q.includes('title case') ||
    q.includes('titlecase') ||
    q.includes('uppercase') ||
    q.includes('lowercase') ||
    q.includes('capitalize') ||
    q.includes('caps')
  ) {
    let caseOption: 'none' | 'lower' | 'upper' | 'title' = 'none';
    if (q.includes('title') || q.includes('capitalize')) caseOption = 'title';
    else if (q.includes('upper') || q.includes('caps')) caseOption = 'upper';
    else if (q.includes('lower')) caseOption = 'lower';

    const explicitCol = q.match(/(?:column|col)\s+([a-z])\b/i);
    const targetCol =
      (explicitCol?.[1] ? resolveColumn(explicitCol[1], columns) : null) ||
      resolveColumn(q, columns) ||
      columns.find((c) => !c.isNumeric && !c.isDate) ||
      columns[0]!;

    return {
      message: `I've prepared a text normalization operation for **${targetCol.rawName}** (${targetCol.letter}).\n\n• Action: ${q.includes('trim') ? 'Trim leading/trailing whitespace' : ''} ${caseOption !== 'none' ? `Convert to ${caseOption} case` : ''}\n\nClick **Apply Changes** to clean this column!`,
      proposedAction: {
        name: 'normalize_text',
        args: {
          sheet: currentSheet.name,
          columns: [targetCol.letter],
          trim: true,
          collapseWhitespace: true,
          case: caseOption,
          headerRow: 1,
        },
        explanation: `Normalize text in column ${targetCol.letter} ("${targetCol.rawName}") with trimming and ${caseOption} casing.`,
        category: 'transform',
      },
    };
  }

  // 3. DATE NORMALIZATION
  if (q.includes('date') || q.includes('iso') || q.includes('yyyy-mm-dd') || q.includes('tarikh')) {
    const explicitCol = q.match(/(?:column|col)\s+([a-z])\b/i);
    const targetCol =
      (explicitCol?.[1] ? resolveColumn(explicitCol[1], columns) : null) ||
      columns.find((c) => c.isDate) ||
      resolveColumn(q, columns) ||
      columns.find((c) => c.cleanName.includes('date') || c.cleanName.includes('time')) ||
      columns[0]!;

    let format: 'YYYY-MM-DD' | 'MM/DD/YYYY' | 'DD/MM/YYYY' = 'YYYY-MM-DD';
    // Word boundaries matter: "must", "customer", and "please use" all contain
    // the substring "us" and must not silently select the US date format.
    if (/\bus\b/.test(q) || q.includes('mm/dd/yyyy')) format = 'MM/DD/YYYY';
    else if (/\beu\b/.test(q) || q.includes('dd/mm/yyyy')) format = 'DD/MM/YYYY';

    return {
      message: `I've prepared a date normalization for **${targetCol.rawName}** (Column ${targetCol.letter}).\n\n• Target Format: **${format}**\n• Total rows: ${dataRowsCount}\n\nClick **Apply Changes** below to standardize all dates!`,
      proposedAction: {
        name: 'format_dates',
        args: {
          sheet: currentSheet.name,
          column: targetCol.letter,
          format,
          headerRow: 1,
        },
        explanation: `Standardize dates in column ${targetCol.letter} ("${targetCol.rawName}") to ${format}.`,
        category: 'format',
      },
    };
  }

  // 4. SORTING
  if (
    q.includes('sort') ||
    q.includes('order by') ||
    q.includes('order the') ||
    q.includes('arrange') ||
    q.includes('ascending') ||
    q.includes('descending')
  ) {
    const isDesc =
      q.includes('desc') ||
      q.includes('z-a') ||
      q.includes('highest') ||
      q.includes('largest') ||
      q.includes('latest');
    const targetCol = resolveColumn(q, columns) || columns[0]!;

    return {
      message: `I've set up a sort operation on **${currentSheet.name}**:\n\n• Target Column: **${targetCol.rawName}** (${targetCol.letter})\n• Direction: **${isDesc ? 'Descending (Z-A / High to Low)' : 'Ascending (A-Z / Low to High)'}**\n\nClick **Apply Changes** below to reorder the spreadsheet.`,
      proposedAction: {
        name: 'sort_range',
        args: {
          sheet: currentSheet.name,
          column: targetCol.letter,
          direction: isDesc ? 'desc' : 'asc',
          startRow: 2,
          startColumn: 'A',
        },
        explanation: `Sort rows by column ${targetCol.letter} ("${targetCol.rawName}") in ${isDesc ? 'descending' : 'ascending'} order.`,
        category: 'transform',
      },
    };
  }

  // 5. MATH CALCULATIONS & AGGREGATIONS (SUM, AVG, MIN, MAX, COUNT)
  if (
    q.includes('sum') ||
    q.includes('total') ||
    q.includes('average') ||
    q.includes('avg') ||
    q.includes('mean') ||
    q.includes('kul') ||
    q.includes('maximum') ||
    q.includes('max') ||
    q.includes('highest') ||
    q.includes('minimum') ||
    q.includes('min') ||
    q.includes('lowest')
  ) {
    const targetCol = resolveColumn(q, columns) || columns.find((c) => c.isNumeric);
    if (targetCol && targetCol.numericValues.length > 0) {
      const sum = targetCol.sum ?? targetCol.numericValues.reduce((a, b) => a + b, 0);
      const avg = targetCol.avg ?? sum / targetCol.numericValues.length;
      const min = targetCol.min ?? Math.min(...targetCol.numericValues);
      const max = targetCol.max ?? Math.max(...targetCol.numericValues);
      const count = targetCol.numericValues.length;

      return {
        message: `**Mathematical & Statistical Analysis for "${targetCol.rawName}" (Column ${targetCol.letter})**:\n\n• **Total Sum:** ${sum.toLocaleString(undefined, { maximumFractionDigits: 2 })}\n• **Average (Mean):** ${avg.toLocaleString(undefined, { maximumFractionDigits: 2 })}\n• **Highest Value (Max):** ${max.toLocaleString()}\n• **Lowest Value (Min):** ${min.toLocaleString()}\n• **Numeric Count:** ${count} non-empty records (out of ${dataRowsCount} rows)\n\nWould you like me to sort by this column, filter values above average, or add a calculated summary column?`,
      };
    }
  }

  // 6. CATEGORY / FREQUENCY BREAKDOWN (PIVOT-LIKE INTELLIGENCE)
  if (
    q.includes('who') ||
    q.includes('breakdown') ||
    q.includes('distribution') ||
    q.includes('group by') ||
    q.includes('most') ||
    q.includes('kaun') ||
    q.includes('kiska') ||
    q.includes('share') ||
    q.includes('how many in vs out')
  ) {
    const targetCol =
      resolveColumn(q, columns) ||
      columns.find((c) => c.distinct.size > 1 && c.distinct.size <= 20) ||
      columns[0]!;
    const sortedEntries = Array.from(targetCol.distinct.entries()).sort((a, b) => b[1] - a[1]);
    const topEntries = sortedEntries.slice(0, 6);

    const breakdownLines = topEntries.map(([val, cnt]) => {
      const pct = ((cnt / Math.max(1, targetCol.nonBlankCount)) * 100).toFixed(1);
      return `• **${val}:** ${cnt} records (${pct}%)`;
    });

    const topItem = sortedEntries[0];

    return {
      message: `**Distribution Breakdown for "${targetCol.rawName}" (Column ${targetCol.letter})**:\n\n${breakdownLines.join('\n')}\n\n**Top Category:** **${topItem?.[0] ?? 'N/A'}** with ${topItem?.[1] ?? 0} transactions (${(((topItem?.[1] ?? 0) / Math.max(1, targetCol.nonBlankCount)) * 100).toFixed(1)}%).\n\nWould you like to filter the sheet to only show any of these categories?`,
    };
  }

  // 7. VALUE-BASED SEARCH / FILTER / "LIST OUT" (e.g. "list out the stocks having 8 items", "me the stock having 8", "filter stock 8")
  // Check if query contains a number or specific value
  const numInQuery = q.match(/\b(\d+(?:\.\d+)?)\b/);
  const operatorGuess: 'equals' | 'gt' | 'lt' | 'gte' | 'lte' =
    q.includes('greater') ||
    q.includes('more than') ||
    q.includes('above') ||
    q.includes('>') ||
    q.includes('jyada')
      ? 'gt'
      : q.includes('less') ||
          q.includes('fewer') ||
          q.includes('below') ||
          q.includes('<') ||
          q.includes('kam')
        ? 'lt'
        : 'equals';

  // Check if query matches a column and either has a number or a categorical value
  const targetColForFilter = resolveColumn(q, columns);
  if (targetColForFilter) {
    let targetValue: string | number | undefined = undefined;

    if (numInQuery && numInQuery[1]) {
      targetValue = Number(numInQuery[1]);
    } else {
      // Check if any distinct value in this column is mentioned in the query
      for (const [distinctVal] of targetColForFilter.distinct) {
        if (q.includes(distinctVal.toLowerCase())) {
          targetValue = distinctVal;
          break;
        }
      }
    }

    if (targetValue !== undefined) {
      // Find matching rows in sheet
      const matchingRowIndices: number[] = [];
      const sampleMatches: string[] = [];

      for (let r = 1; r < currentSheet.rows.length; r++) {
        const row = currentSheet.rows[r];
        const cell = row?.[targetColForFilter.index];
        const rawVal = cell?.value;
        const numVal =
          typeof rawVal === 'number'
            ? rawVal
            : Number(
                String(rawVal ?? '')
                  .replace(/,/g, '')
                  .trim(),
              );
        const strVal = String(rawVal ?? '')
          .trim()
          .toLowerCase();

        let isMatch = false;
        if (operatorGuess === 'equals') {
          if (typeof targetValue === 'number') {
            isMatch = numVal === targetValue || strVal === String(targetValue);
          } else {
            isMatch =
              strVal === String(targetValue).toLowerCase() ||
              strVal.includes(String(targetValue).toLowerCase());
          }
        } else if (operatorGuess === 'gt' && typeof targetValue === 'number') {
          isMatch = !isNaN(numVal) && numVal > targetValue;
        } else if (operatorGuess === 'lt' && typeof targetValue === 'number') {
          isMatch = !isNaN(numVal) && numVal < targetValue;
        }

        if (isMatch) {
          matchingRowIndices.push(r + 1); // 1-indexed for display
          if (sampleMatches.length < 5) {
            // Build a quick summary of this row
            const details: string[] = [];
            for (let c = 0; c < Math.min(6, columns.length); c++) {
              if (
                c !== targetColForFilter.index &&
                row?.[c]?.value !== null &&
                row?.[c]?.value !== undefined
              ) {
                const hName = columns[c]?.rawName || `Col ${indexToColumn(c)}`;
                details.push(`${hName}: \`${row[c]?.value}\``);
              }
            }
            sampleMatches.push(`• **Row ${r + 1}:** ${details.slice(0, 3).join(' | ')}`);
          }
        }
      }

      const count = matchingRowIndices.length;
      if (count > 0) {
        return {
          message: `I analyzed all **${dataRowsCount} rows** in **${currentSheet.name}**.\n\nFound **${count} matching record(s)** where **${targetColForFilter.rawName}** is **${targetValue}**:\n\n${sampleMatches.join('\n')}${count > 5 ? `\n• *...and ${count - 5} more rows (${matchingRowIndices.slice(5).join(', ')})*` : ''}\n\nI've prepared a **filter operation** to isolate these ${count} rows on your grid. Review the preview card below and click **Apply Changes** to filter the view!`,
          proposedAction: {
            name: 'filter_rows',
            args: {
              sheet: currentSheet.name,
              column: targetColForFilter.letter,
              operator: operatorGuess,
              value: targetValue,
              headerRow: 1,
            },
            explanation: `Filter rows where ${targetColForFilter.rawName} (${targetColForFilter.letter}) ${operatorGuess} "${targetValue}" (${count} matching rows).`,
            category: 'filter',
          },
        };
      } else {
        const samples = Array.from(targetColForFilter.distinct.keys()).slice(0, 5);
        return {
          message: `I searched column **${targetColForFilter.rawName}** (Column ${targetColForFilter.letter}) across all **${dataRowsCount} rows**, but found **0 records** matching **"${targetValue}"**.\n\nExisting sample values in this column are: ${samples.map((s) => `\`${s}\``).join(', ')}.`,
        };
      }
    }
  }

  // 8. MISSING DATA & AUDITING
  if (
    q.includes('blank') ||
    q.includes('empty') ||
    q.includes('missing') ||
    q.includes('null') ||
    q.includes('khali')
  ) {
    const blankReports: string[] = [];
    for (const col of columns) {
      const blanks = dataRowsCount - col.nonBlankCount;
      if (blanks > 0) {
        const pct = ((blanks / Math.max(1, dataRowsCount)) * 100).toFixed(1);
        blankReports.push(
          `• **${col.rawName}** (${col.letter}): ${blanks} missing cells (${pct}%)`,
        );
      }
    }

    if (blankReports.length > 0) {
      return {
        message: `**Missing Value Audit for "${currentSheet.name}"** (${dataRowsCount} data rows):\n\n${blankReports.join('\n')}\n\nAll other columns are 100% complete. Would you like to filter out empty rows or fill missing cells with a default value?`,
      };
    } else {
      return {
        message: `Audit Complete: All ${columns.length} columns in **${currentSheet.name}** are complete with zero missing or blank cells across all ${dataRowsCount} rows.`,
      };
    }
  }

  // 9. GENERAL SUMMARY & OVERVIEW
  if (
    q.includes('summary') ||
    q.includes('overview') ||
    q.includes('tell me') ||
    q.includes('what is') ||
    q.includes('kya hai')
  ) {
    const colSummary = columns
      .map(
        (c) =>
          `• **${c.rawName}** (${c.letter}): ${c.isNumeric ? 'Numeric' : c.isDate ? 'Date' : 'Text'} (${c.distinct.size} unique values)`,
      )
      .slice(0, 8);

    return {
      message: `**Worksheet Overview: "${currentSheet.name}"**\n\n• **Total Records:** ${dataRowsCount} data rows\n• **Total Columns:** ${columns.length} columns\n• **Worksheets in Workbook:** ${workbook.sheets.map((s) => s.name).join(', ')}\n\n**Column Profiles:**\n${colSummary.join('\n')}${columns.length > 8 ? `\n• *...and ${columns.length - 8} more columns*` : ''}\n\nWhat would you like me to do? You can ask to filter rows, calculate sums/averages, remove duplicates, sort, or standardize dates.`,
    };
  }

  // 10. FALLBACK SMART REASONING: Search sheet cells
  const searchMatches = searchCellsInSheet(currentSheet, userQuery);
  if (searchMatches.length > 0) {
    const first = searchMatches[0]!;
    return {
      message: `I found **${searchMatches.length} matching cell(s)** for "${userQuery}" across **${currentSheet.name}** (e.g. Cell ${first.column}${first.row}: "${first.text}").\n\nI've prepared a filter operation to show only rows containing this text. Click **Apply Changes** below to view them!`,
      proposedAction: {
        name: 'filter_rows',
        args: {
          sheet: currentSheet.name,
          column: first.column,
          operator: 'contains',
          value: userQuery.trim(),
          headerRow: 1,
        },
        explanation: `Filter rows where column ${first.column} contains "${userQuery.trim()}".`,
        category: 'filter',
      },
    };
  }

  // FINAL DEFAULT HELPFUL SUGGESTION
  const sampleHeaders = columns.map((c) => c.rawName).slice(0, 4);
  return {
    message: `I analyzed your request: "${userQuery}".\n\nI am ready to perform high-speed operations on **${currentSheet.name}** (${dataRowsCount} rows). Here are common things I can do for you:\n• *"List out rows where ${sampleHeaders[sampleHeaders.length - 1] || 'stock'} is 8"*\n• *"Total sum of ${columns.find((c) => c.isNumeric)?.rawName || 'Quantity'}"*\n• *"Remove all duplicate rows"*\n• *"Sort by ${sampleHeaders[0] || 'Column A'} descending"*\n• *"Who handled the most transactions?"*\n• *"Find missing values"*`,
  };
}

export function parseNaturalLanguageIntent(
  userQuery: string,
  workbook: Workbook,
  activeSheetName: string,
): ProposedAction | null {
  const result = analyzeSpreadsheetIntentAndData(userQuery, workbook, activeSheetName);
  return result.proposedAction || null;
}
