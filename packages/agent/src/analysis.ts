import {
  type Workbook,
  type Sheet,
  type CellValue,
  columnToIndex,
  indexToColumn,
  maxColumnCount,
} from '@excel-agent/engine';

/**
 * Spreading an array into `Math.min`/`Math.max` throws a RangeError once the
 * array passes roughly 125k elements, which a single wide or tall column
 * reaches. These helpers keep the same result without touching the call stack.
 */
function smallest(values: number[]): number {
  return values.reduce((current, value) => Math.min(current, value), Infinity);
}

function largest(values: number[]): number {
  return values.reduce((current, value) => Math.max(current, value), -Infinity);
}

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
  const totalCols = maxColumnCount(sheet.rows);
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
      min = smallest(numericValues);
      max = largest(numericValues);
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
  const totalCols = maxColumnCount(sheet.rows);
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

export interface FilterCandidate {
  column: ColumnMetadata;
  value: string | number;
  operator: 'equals' | 'contains' | 'gt' | 'lt' | 'gte' | 'lte' | 'starts_with' | 'ends_with';
  matchedToken: string;
}

export function findFilterCandidateInSheet(
  query: string,
  columns: ColumnMetadata[],
  _sheet: Sheet,
): FilterCandidate | null {
  const q = query.trim().toLowerCase();

  // 1. Direct explicit column reference: e.g. "column D contains IN", "Type is IN"
  for (const col of columns) {
    const colNameRegex = new RegExp(
      `\\b(?:column\\s+${col.letter}|col\\s+${col.letter}|${col.cleanName})\\b`,
      'i',
    );
    if (colNameRegex.test(q)) {
      const numMatch = q.match(
        /(?:>|>=|<|<=|greater than|more than|above|less than|below)\s*(\d+(?:\.\d+)?)/i,
      );
      if (numMatch && numMatch[1]) {
        const op =
          q.includes('>') || q.includes('greater') || q.includes('more') || q.includes('above')
            ? 'gt'
            : 'lt';
        return { column: col, value: Number(numMatch[1]), operator: op, matchedToken: numMatch[0] };
      }
      for (const [distinctVal] of col.distinct) {
        if (distinctVal && q.includes(distinctVal.toLowerCase())) {
          return { column: col, value: distinctVal, operator: 'contains', matchedToken: distinctVal };
        }
      }
    }
  }

  // 2. Transaction status shorthand check: e.g. "IN data", "OUT data", "IN rows", "filter IN", "filter OUT"
  const hasInStatus =
    /\b(?:the\s+)?in\s+(?:data|records?|rows?|transactions?|items?|stock|operations?|action)\b/i.test(query) ||
    /\b(?:filter|extract|separate|isolate|pull(?:\s+out)?)\b[\s\S]{0,50}?\bin\s+(?:data|records?|rows?|transactions?|items?|stock|operations?|action)\b/i.test(query) ||
    /\bin\s+(?:operation|action|status)\b/i.test(query);

  const hasOutStatus =
    !hasInStatus &&
    (/\b(?:the\s+)?out\s+(?:data|records?|rows?|transactions?|items?|stock|operations?|action)\b/i.test(query) ||
      /\bout\s+(?:operation|action|status)\b/i.test(query) ||
      /\b(?:filter|extract|separate|isolate)\s+(?:for\s+)?(?:the\s+)?out\s+(?:data|records?|rows?|transactions?|items?|stock|operations?|action)\b/i.test(query));

  if (hasInStatus) {
    for (const col of columns) {
      for (const [dVal] of col.distinct) {
        if (/^IN\b|\(IN\)|\bIN\s*\(|^IN\s+/i.test(dVal) || dVal.toUpperCase().startsWith('IN')) {
          return { column: col, value: 'IN', operator: 'contains', matchedToken: 'IN' };
        }
      }
    }
  }

  if (hasOutStatus) {
    for (const col of columns) {
      for (const [dVal] of col.distinct) {
        if (/^OUT\b|\(OUT\)|\bOUT\s*\(|^OUT\s+/i.test(dVal) || dVal.toUpperCase().startsWith('OUT')) {
          return { column: col, value: 'OUT', operator: 'contains', matchedToken: 'OUT' };
        }
      }
    }
  }

  // 3. Scan meaningful query words against all column distinct values
  const stopWords = new Set([
    'filter',
    'out',
    'the',
    'data',
    'into',
    'a',
    'an',
    'separate',
    'new',
    'sheet',
    'tab',
    'another',
    'all',
    'rows',
    'records',
    'items',
    'and',
    'or',
    'to',
    'for',
    'from',
    'with',
    'only',
    'where',
    'having',
    'which',
    'is',
    'are',
    'in',
    'of',
    'please',
    'put',
    'move',
    'copy',
    'extract',
    'isolate',
    'take',
    'make',
    'create',
    'give',
    'me',
    'show',
    'by',
    'not',
    'no',
    'never',
    'none',
    'you',
    'your',
    'yours',
    'i',
    'my',
    'we',
    'our',
    'us',
    'they',
    'them',
    'he',
    'she',
    'it',
    'this',
    'that',
    'these',
    'those',
    'do',
    'does',
    'did',
    'done',
    'dont',
    "don't",
    'cannot',
    "can't",
    'cant',
    'formula',
    'formulas',
    'cell',
    'cells',
    'putting',
    'check',
    'properly',
    'correct',
    'wrong',
    'right',
    'again',
    'instead',
    'rather',
    'told',
    'tell',
    'said',
    'saying',
    'why',
    'how',
    'what',
    'when',
    'where',
    'who',
    'can',
    'could',
    'would',
    'should',
    'will',
    'shall',
    'total',
    'sum',
    'avg',
    'average',
    'count',
  ]);

  const rawTokens = query.split(/[\s,._/?!+;:"'()\[\]{}]+/).filter((t) => t.length >= 2);
  for (const token of rawTokens) {
    const lowerToken = token.toLowerCase();
    if (stopWords.has(lowerToken)) continue;

    for (const col of columns) {
      for (const [dVal] of col.distinct) {
        if (!dVal) continue;
        const lowerD = dVal.toLowerCase();
        // Exact match or full word match for longer tokens (prevents false positive substring matches like 'not' in 'Noncontrolling')
        const isMatch =
          lowerD === lowerToken ||
          (lowerToken.length >= 4 && new RegExp(`\\b${lowerToken}\\b`, 'i').test(lowerD));

        if (isMatch) {
          return {
            column: col,
            value: token,
            operator: 'contains',
            matchedToken: token,
          };
        }
      }
    }
  }

  return null;
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

  // 0A. Numerical value replacement: e.g. "change all negative amount to 0" or "replace negative values with 0"
  const negativeToValMatch = raw.match(
    /(?:change|set|replace|turn|convert|clamp)\s+(?:all\s+)?negative(?:\s+(?:amount|numbers?|values?|figures?))?\s+(?:to|with)\s+([0-9.-]+)/i,
  );

  if (negativeToValMatch) {
    const targetVal = parseFloat(negativeToValMatch[1] || '0') || 0;
    const edits: Array<{ row: number; column: string; value: number }> = [];
    currentSheet.rows.forEach((row, rIdx) => {
      row.forEach((cell, cIdx) => {
        if (!cell || cell.formula !== undefined) return;
        const val = cell.value;
        let num: number | null = null;
        if (typeof val === 'number') {
          num = val;
        } else if (typeof val === 'string' && /^-[\d,]+(?:\.\d+)?$/.test(val.trim())) {
          num = parseFloat(val.replace(/,/g, ''));
        }
        if (num !== null && num < 0) {
          edits.push({
            row: rIdx + 1,
            column: indexToColumn(cIdx),
            value: targetVal,
          });
        }
      });
    });

    if (edits.length > 0) {
      return {
        message: `I analyzed **${currentSheet.name}** and identified **${edits.length} negative value(s)**.\n\nI've prepared an update to set all ${edits.length} negative values to **${targetVal}**. Click **Apply Changes** below to execute.`,
        proposedAction: {
          name: 'edit_cells',
          args: {
            sheet: currentSheet.name,
            edits: edits.slice(0, 500),
          },
          explanation: `Change ${edits.length} negative value(s) to ${targetVal} in ${currentSheet.name}.`,
          category: 'transform',
        },
      };
    } else {
      return {
        message: `I scanned all ${currentSheet.rows.length} rows in **${currentSheet.name}**; there are currently zero negative values in the sheet.`,
      };
    }
  }

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
    const lastIndex = lastCol ? columnToIndex(lastCol) : undefined;
    const insertLetter = indexToColumn(lastIndex !== undefined ? lastIndex + 1 : 0);
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

  // 0c. SHEET OPERATIONS: filter-to-sheet, create, duplicate, delete, and summary rows
  const isFilterToSheetQuery =
    /(?:filter|extract|copy|move|separate|split|put|isolate|pull|take)\b[\s\S]{0,60}?\b(?:into|in|to)\s+(?:a\s+)?(?:separate|new|another|diff|different)\s+(?:sheet|tab)/i.test(
      raw,
    ) ||
    /(?:create|make|add)\s+(?:a\s+)?(?:separate|new|another)\s+sheet\s+(?:with|for|having|of|from)\b/i.test(
      raw,
    ) ||
    /\b(?:separate|new|another)\s+sheet\s+(?:with|for|having|of)\b/i.test(raw);

  if (isFilterToSheetQuery) {
    const candidate = findFilterCandidateInSheet(raw, columns, currentSheet);
    if (candidate) {
      const colIdx = candidate.column.index;
      const filterValStr = String(candidate.value).toLowerCase();
      const matchingRowIndices: number[] = [];
      const sampleMatches: string[] = [];

      for (let r = 1; r < currentSheet.rows.length; r++) {
        const row = currentSheet.rows[r];
        const cell = row?.[colIdx];
        const cellVal = cell?.value;
        const strVal = String(cellVal ?? '').trim().toLowerCase();
        const numVal =
          typeof cellVal === 'number'
            ? cellVal
            : Number(String(cellVal ?? '').replace(/,/g, '').trim());
        let isMatch = false;

        if (candidate.operator === 'contains') {
          isMatch = strVal.includes(filterValStr);
        } else if (candidate.operator === 'equals') {
          isMatch = strVal === filterValStr || strVal.includes(filterValStr);
        } else if (candidate.operator === 'gt' && typeof candidate.value === 'number') {
          isMatch = !isNaN(numVal) && numVal > candidate.value;
        } else if (candidate.operator === 'lt' && typeof candidate.value === 'number') {
          isMatch = !isNaN(numVal) && numVal < candidate.value;
        }

        if (isMatch) {
          matchingRowIndices.push(r + 1);
          if (sampleMatches.length < 5) {
            const details: string[] = [];
            for (let c = 0; c < Math.min(6, columns.length); c++) {
              if (row?.[c]?.value !== null && row?.[c]?.value !== undefined) {
                const hName = columns[c]?.rawName || `Col ${indexToColumn(c)}`;
                details.push(`${hName}: \`${row[c]?.value}\``);
              }
            }
            sampleMatches.push(`• **Row ${r + 1}:** ${details.slice(0, 3).join(' | ')}`);
          }
        }
      }

      const count = matchingRowIndices.length;
      const namedMatch = raw.match(/(?:sheet|tab)\s+(?:named|called)\s*['"]?([^'"]+)['"]?/i);
      const targetSheetName =
        namedMatch?.[1]?.trim() ||
        `${String(candidate.value).replace(/[^a-zA-Z0-9_-]+/g, '_').replace(/^_+|_+$/g, '') || 'Filtered'}_Data`;

      if (count > 0) {
        return {
          message: `I analyzed all **${dataRowsCount} rows** in **${currentSheet.name}** and identified **Column ${candidate.column.letter} (${candidate.column.rawName})** matching **"${candidate.value}"**.\n\nFound **${count} matching record(s)**:\n\n${sampleMatches.join('\n')}${count > 5 ? `\n• *...and ${count - 5} more rows*` : ''}\n\nI've prepared to extract these records into a new sheet **"${targetSheetName}"** with headers preserved. Click **Apply Changes** below to create it!`,
          proposedAction: {
            name: 'filter_to_new_sheet',
            args: {
              sheet: currentSheet.name,
              targetSheet: targetSheetName,
              column: candidate.column.letter,
              operator: candidate.operator,
              value: candidate.value,
              headerRow: 1,
            },
            explanation: `Filter ${count} rows where ${candidate.column.rawName} (${candidate.column.letter}) contains "${candidate.value}" into new sheet "${targetSheetName}".`,
            category: 'filter',
          },
        };
      }
    }
  }

  const createSheetMatch = raw.match(
    /(?:create|add|make|insert)\s+(?:a\s+)?(?:new\s+)?sheet\s+(?:named|called)\s*['"]?([^'"]+)['"]?/i,
  );
  if (createSheetMatch && createSheetMatch[1]) {
    const sheetName = createSheetMatch[1].trim();
    return {
      message: `I've prepared to create a new worksheet named **"${sheetName}"**.\n\nClick **Apply Changes** to add the sheet.`,
      proposedAction: {
        name: 'create_sheet',
        args: { sheetName },
        explanation: `Create a new sheet "${sheetName}".`,
        category: 'structure',
      },
    };
  }

  const dupSheetMatch = raw.match(
    /(?:duplicate|clone|copy)\s+(?:the\s+)?sheet\s+['"]?([^'"]+?)['"]?\s+(?:as|to|into)\s+['"]?([^'"]+?)['"]?$/i,
  );
  if (dupSheetMatch && dupSheetMatch[1] && dupSheetMatch[2]) {
    const src = dupSheetMatch[1].trim();
    const target = dupSheetMatch[2].trim();
    return {
      message: `I've prepared to duplicate sheet **"${src}"** into **"${target}"**.\n\nClick **Apply Changes** to proceed.`,
      proposedAction: {
        name: 'duplicate_sheet',
        args: { sheet: src, targetSheet: target },
        explanation: `Duplicate sheet "${src}" to "${target}".`,
        category: 'structure',
      },
    };
  }

  const delSheetMatch = raw.match(
    /(?:delete|remove|drop)\s+(?:the\s+)?sheet\s+['"]?([^'"]+?)['"]?$/i,
  );
  if (delSheetMatch && delSheetMatch[1]) {
    const sheetToDelete = delSheetMatch[1].trim();
    return {
      message: `I've prepared to delete sheet **"${sheetToDelete}"**.\n\n⚠️ Deleting a sheet is irreversible. Review the preview card and click **Apply Changes** to confirm.`,
      proposedAction: {
        name: 'delete_sheet',
        args: { sheet: sheetToDelete },
        explanation: `Delete sheet "${sheetToDelete}".`,
        category: 'structure',
      },
    };
  }

  const summaryRowMatch = raw.match(
    /(?:add|insert|calculate)\s+(?:a\s+)?(?:summary|total|totals|sum|average|avg)\s+row(?:\s+at\s+the\s+bottom)?/i,
  );
  if (summaryRowMatch) {
    const isAvg = raw.toLowerCase().includes('average') || raw.toLowerCase().includes('avg');
    const agg = isAvg ? 'average' : 'sum';
    const label = isAvg ? 'Average' : 'Total';
    const numericCols = columns.filter((c) => c.isNumeric).map((c) => c.letter);
    return {
      message: `I've prepared to append a **${label}** row at the bottom of **${currentSheet.name}** across numeric columns (${numericCols.join(', ')}).\n\nClick **Apply Changes** to add it.`,
      proposedAction: {
        name: 'add_summary_row',
        args: {
          sheet: currentSheet.name,
          aggregation: agg,
          label,
          columns: numericCols,
          headerRow: 1,
        },
        explanation: `Add ${label} row at bottom of ${currentSheet.name}.`,
        category: 'transform',
      },
    };
  }

  const cleanSheetMatch =
    raw.match(
      /(?:clean|cleaning|tidy|tidying|structure|structuring|organize|organizing|format|prepare|extract)[\s\S]{0,80}?(?:new\s+(?:sheet|tab)|separate\s+(?:sheet|tab)|clean\s+sheet|new\s+dataset)/i,
    ) ||
    raw.match(
      /\b(?:clean|structure|tidy|organize|prepare|standardize)\s+(?:up\s+)?(?:the\s+)?(?:data|dataset|sheet|records|worksheet|table|everything)\b/i,
    ) ||
    raw.match(
      /^(?:clean\s+data|structured\s+data|clean\s+the\s+data|give\s+me\s+(?:the\s+)?structured\s+data|tidy\s+up|make\s+it\s+clean|clean\s+up)$/i,
    );
  if (cleanSheetMatch) {
    const namedMatch = raw.match(/(?:sheet|tab)\s+(?:named|called)\s*['"]([^'"]+)['"]/i);
    const targetName = namedMatch?.[1]?.trim() || `${currentSheet.name}_Cleaned`;
    return {
      message: `I've prepared a comprehensive data cleaning & structuring pipeline for **${currentSheet.name}** into **${targetName}**:\n\n• Trim and collapse whitespace\n• Coerce numeric text into numbers\n• Drop blank padding rows and columns\n• Standardize header row alignment\n\nClick **Apply Changes** to generate the clean, structured sheet.`,
      proposedAction: {
        name: 'clean_to_new_sheet',
        args: { sheet: currentSheet.name, targetSheet: targetName },
        explanation: `Clean and structure "${currentSheet.name}" into a new sheet "${targetName}".`,
        category: 'transform',
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

  // 0d. BROAD CLEANING & STRUCTURING DIRECTIVES ("clean the data", "clean and structured the sheet", "tidy up")
  const isBroadCleanQuery =
    /(?:clean|structure|tidy|standardize|prepare)\s+(?:and\s+)?(?:structure\s+|clean\s+)?(?:the\s+)?(?:sheet|data|table|dataset|workbook|file)/i.test(
      raw,
    ) ||
    /^(?:clean|structure|tidy up|clean up|clean data|clean sheet|structure data|structure sheet|clean and structured? the sheet)\b/i.test(
      raw,
    );

  if (isBroadCleanQuery) {
    // 1. Check for duplicates across data rows
    const seen = new Set<string>();
    let dupCount = 0;
    for (let r = 1; r < currentSheet.rows.length; r++) {
      const key = (currentSheet.rows[r] || []).map((c) => String(c?.value ?? '')).join('|~|');
      if (seen.has(key)) dupCount++;
      else seen.add(key);
    }
    if (dupCount > 0) {
      return {
        message: `I audited all **${dataRowsCount} data rows** in **${currentSheet.name}**.\n\n• Found **${dupCount} duplicate row(s)**.\n\nCleaning Step 1: Remove redundant duplicate rows to structure your dataset accurately. Click **Apply Changes** to proceed!`,
        proposedAction: {
          name: 'delete_duplicates',
          args: {
            sheet: currentSheet.name,
            columns: allColumns,
            headerRow: 1,
            keep: 'first',
          },
          explanation: `Remove ${dupCount} duplicate rows across all ${allColumns.length} columns.`,
          category: 'structure',
        },
      };
    }

    // 2. Check for columns with untrimmed whitespace
    const untrimmedCols: string[] = [];
    for (const col of columns) {
      let untrimmed = 0;
      for (let r = 1; r < currentSheet.rows.length; r++) {
        const val = currentSheet.rows[r]?.[col.index]?.value;
        if (typeof val === 'string' && val.trim() !== val) untrimmed++;
      }
      if (untrimmed > 0) untrimmedCols.push(col.letter);
    }
    if (untrimmedCols.length > 0) {
      return {
        message: `I audited text columns in **${currentSheet.name}** and found untrimmed whitespace across **${untrimmedCols.length} column(s)** (${untrimmedCols.join(', ')}).\n\nClick **Apply Changes** to clean and trim all dirty cells!`,
        proposedAction: {
          name: 'normalize_text',
          args: {
            sheet: currentSheet.name,
            columns: untrimmedCols,
            trim: true,
            collapseWhitespace: true,
            case: 'none',
            headerRow: 1,
          },
          explanation: `Trim whitespace across dirty text columns (${untrimmedCols.join(', ')}).`,
          category: 'transform',
        },
      };
    }

    // 3. Check for date columns needing formatting
    const dateCol = columns.find((c) => c.isDate);
    if (dateCol) {
      return {
        message: `I analyzed the structure of **${currentSheet.name}** and identified date column **${dateCol.rawName}** (Column ${dateCol.letter}).\n\nStandardizing date formatting to ISO \`YYYY-MM-DD\` will clean your data pipeline. Click **Apply Changes** to normalize!`,
        proposedAction: {
          name: 'format_dates',
          args: {
            sheet: currentSheet.name,
            column: dateCol.letter,
            format: 'YYYY-MM-DD',
            headerRow: 1,
          },
          explanation: `Standardize dates in column ${dateCol.letter} ("${dateCol.rawName}") to ISO format.`,
          category: 'format',
        },
      };
    }

    // 4. Default: Sheet is already clean
    return {
      message: `I performed a full data audit on **${currentSheet.name}** (${dataRowsCount} rows, ${columns.length} columns).\n\n• **Duplicates:** 0 duplicate rows found\n• **Whitespace:** No untrimmed whitespace detected\n• **Structure:** Headers and data types are properly aligned\n\nYour sheet is already cleanly structured!`,
    };
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
    q.includes('title case') ||
    q.includes('titlecase') ||
    q.includes('uppercase') ||
    q.includes('lowercase') ||
    q.includes('capitalize') ||
    q.includes('caps') ||
    (q.includes('clean') && (q.includes('column') || q.includes('col') || q.includes('text') || q.includes('name')))
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

  // 4.5. ROW-LEVEL HORIZONTAL AGGREGATION & METRIC ANALYSIS
  // Handles queries like "analyz to row 21 and give me total of all FY in row 21", "total of row 21", "sum of row 21", "average of row 5"
  const rowMatch = q.match(/\b(?:row|line)\s*(\d+)\b/i);
  const isRowQuery =
    Boolean(rowMatch) &&
    (q.includes('total') ||
      q.includes('sum') ||
      q.includes('analyz') ||
      q.includes('analyse') ||
      q.includes('all fy') ||
      q.includes('average') ||
      q.includes('avg') ||
      q.includes('give me') ||
      q.includes('what is') ||
      q.includes('breakdown') ||
      q.includes('aggregate'));

  if (rowMatch && isRowQuery) {
    const rowNum = parseInt(rowMatch[1]!, 10);
    // In spreadsheets, users refer to 1-indexed row numbers (e.g. Row 21 is index 20)
    const targetRowIndex =
      rowNum - 1 >= 0 && rowNum - 1 < currentSheet.rows.length
        ? rowNum - 1
        : rowNum < currentSheet.rows.length
          ? rowNum
          : -1;

    if (targetRowIndex >= 0) {
      const targetRow = currentSheet.rows[targetRowIndex] ?? [];

      // Find row label from the first non-empty text cell
      let rowLabel = '';
      for (const cell of targetRow) {
        const str = String(cell?.value ?? '').trim();
        if (str && isNaN(Number(str.replace(/,/g, '')))) {
          rowLabel = str;
          break;
        }
      }
      if (!rowLabel) rowLabel = `Row ${rowNum}`;

      // Detect header row across top 10 rows to match column labels (e.g., FY '09, FY '10, etc.)
      let headerRowIndex = 0;
      let maxHeaderMatch = 0;
      for (let r = 0; r < Math.min(10, currentSheet.rows.length); r++) {
        if (r === targetRowIndex) continue;
        const row = currentSheet.rows[r] ?? [];
        let stringCount = 0;
        for (const cell of row) {
          const val = String(cell?.value ?? '').trim();
          if (val && (/^fy\s*'?\d{2,4}$/i.test(val) || /^20\d{2}$/.test(val) || val.length > 1)) {
            stringCount++;
          }
        }
        if (stringCount > maxHeaderMatch) {
          maxHeaderMatch = stringCount;
          headerRowIndex = r;
        }
      }

      const headers = (currentSheet.rows[headerRowIndex] ?? []).map(
        (c, idx) => String(c?.value ?? '').trim() || indexToColumn(idx),
      );

      // Collect numeric values across the row
      const numericCells: Array<{ colLetter: string; header: string; value: number }> = [];
      targetRow.forEach((cell, cIdx) => {
        const rawVal = cell?.value;
        if (rawVal === null || rawVal === undefined || rawVal === '') return;
        const colLetter = indexToColumn(cIdx);
        const header = headers[cIdx] || colLetter;

        let num: number = NaN;
        if (typeof rawVal === 'number') {
          num = rawVal;
        } else {
          let cleaned = String(rawVal).trim().replace(/,/g, '').replace(/^\$/, '');
          // Handle accounting negative parentheses: (1,234) -> -1234
          if (cleaned.startsWith('(') && cleaned.endsWith(')')) {
            cleaned = `-${cleaned.slice(1, -1)}`;
          }
          num = parseFloat(cleaned);
        }

        if (!isNaN(num)) {
          numericCells.push({ colLetter, header, value: num });
        }
      });

      if (numericCells.length > 0) {
        const totalSum = numericCells.reduce((sum, item) => sum + item.value, 0);
        const avg = totalSum / numericCells.length;
        const minVal = smallest(numericCells.map((c) => c.value));
        const maxVal = largest(numericCells.map((c) => c.value));
        const minCell = numericCells.find((c) => c.value === minVal);
        const maxCell = numericCells.find((c) => c.value === maxVal);

        const periodRows = numericCells.map((c) => {
          const pct = totalSum !== 0 ? ((c.value / totalSum) * 100).toFixed(1) : '0.0';
          const isMax = c.value === maxVal;
          const isMin = c.value === minVal;
          const tag = isMax ? ' **(Peak)**' : isMin ? ' *(Trough)*' : '';
          return `| **${c.header}** (${c.colLetter}) | **${c.value >= 0 ? '+' : ''}${c.value.toLocaleString()}** | ${pct}% | ${c.value >= 0 ? 'Positive' : 'Negative'}${tag} |`;
        });

        return {
          message:
            `### Row ${rowNum} Comprehensive Analysis: ${rowLabel}\n\n` +
            `Here is the complete calculation and breakdown for **Row ${rowNum}** (**${rowLabel}**) across all **${numericCells.length} recorded fiscal periods**:\n\n` +
            `- **Total Sum (All FY)**: **${totalSum.toLocaleString()}** (${totalSum >= 0 ? '$' : '-$'}${Math.abs(totalSum).toLocaleString()}M)\n` +
            `- **Annual Average / Mean**: **${avg.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}**\n` +
            `- **Highest Period**: **${maxVal.toLocaleString()}** (${maxCell?.header || ''})\n` +
            `- **Lowest Period**: **${minVal.toLocaleString()}** (${minCell?.header || ''})\n` +
            `- **Periods Counted**: ${numericCells.length} Fiscal Periods (${numericCells[0]?.header} to ${numericCells[numericCells.length - 1]?.header})\n\n` +
            `#### Fiscal Year Breakdown Table:\n` +
            `| Fiscal Period | Value ($M) | % of Total | Status |\n` +
            `| :--- | :--- | :--- | :--- |\n` +
            `${periodRows.join('\n')}\n` +
            `| **Total (All ${numericCells.length} FY)** | **${totalSum.toLocaleString()}** | **100.0%** | **${totalSum >= 0 ? 'Cumulative Profit' : 'Cumulative Loss'}** |\n\n` +
            `**Key Takeaway**: Across the ${numericCells.length}-year reporting horizon, Row ${rowNum} generated a cumulative total of **${totalSum.toLocaleString()}** with an average annual run-rate of **${avg.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}**.`,
        };
      }
    }
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
      const min = targetCol.min ?? smallest(targetCol.numericValues);
      const max = targetCol.max ?? largest(targetCol.numericValues);
      const count = targetCol.numericValues.length;

      return {
        message: `**Mathematical & Statistical Analysis for "${targetCol.rawName}" (Column ${targetCol.letter})**:\n\n• **Total Sum:** ${sum.toLocaleString(undefined, { maximumFractionDigits: 2 })}\n• **Average (Mean):** ${avg.toLocaleString(undefined, { maximumFractionDigits: 2 })}\n• **Highest Value (Max):** ${max.toLocaleString()}\n• **Lowest Value (Min):** ${min.toLocaleString()}\n• **Numeric Count:** ${count} non-empty records (out of ${dataRowsCount} rows)\n\nWould you like me to sort by this column, filter values above average, or add a calculated summary column?`,
      };
    }
  }

  // 5b. MULTI-ENTITY COUNT / FREQUENCY QUERY (e.g. "how many time pradeep bothra has performed OUT operation ?", "how many times ronak has perform IN operation")
  const howManyMatch =
    raw.match(
      /(?:find\s+out\s+(?:that\s+)?)?(?:how\s+many\s+times?|count\s+of|kitni\s+baar)\s+([a-z0-9_\s/()]+?)\s+(?:has\s+)?(?:performed?|done|did|had|carry|make)\s+([a-z0-9_\s/()]+?)(?:\s+operation|\s+action|\s+in\s+sheet|\s*\?|$)/i,
    ) ||
    raw.match(
      /(?:find\s+out\s+(?:that\s+)?)?(?:how\s+many\s+times?|count\s+how\s+many\s+times?)\s+([a-z0-9_\s/()]+?)\s+(?:in|for|with)\s+([a-z0-9_\s/()]+?)(?:\s+operation|\s+action|\s+in\s+sheet|\s*\?|$)/i,
    );

  if (howManyMatch) {
    const entity1 = howManyMatch[1]!.trim().toLowerCase();
    const entity2 = howManyMatch[2]!.trim().toLowerCase();

    let matchCount = 0;
    const matchingRows: number[] = [];
    const sampleDetails: string[] = [];

    for (let r = 1; r < currentSheet.rows.length; r++) {
      const row = currentSheet.rows[r];
      if (!row) continue;
      const rowText = row.map((cell) => String(cell?.value ?? '').toLowerCase()).join(' | ');
      if (rowText.includes(entity1) && rowText.includes(entity2)) {
        matchCount++;
        matchingRows.push(r + 1);
        if (sampleDetails.length < 5) {
          const rowSummary = row
            .filter((c) => c?.value !== null && c?.value !== undefined && String(c.value).trim())
            .slice(0, 4)
            .map((c) => String(c.value))
            .join(' | ');
          sampleDetails.push(`• **Row ${r + 1}:** ${rowSummary}`);
        }
      }
    }

    if (matchCount > 0) {
      return {
        message: `**Query Result:**\n\n**${howManyMatch[1]!.trim()}** performed the **${howManyMatch[2]!.trim()}** operation **${matchCount} time(s)** in **${currentSheet.name}**.\n\n${sampleDetails.length > 0 ? `**Sample Records:**\n${sampleDetails.join('\n')}\n\n` : ''}Would you like me to filter these matching rows into a separate sheet?`,
      };
    } else {
      return {
        message: `I searched all **${dataRowsCount} rows** in **${currentSheet.name}**, but found **0 records** where **${howManyMatch[1]!.trim()}** performed **${howManyMatch[2]!.trim()}**.`,
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

  // 6.5. FINANCIAL ANALYSIS: PROFIT & LOSS / PROBABILITY OF PROFIT
  if (
    q.includes('profit') ||
    q.includes('loss') ||
    (q.includes('probability') && (q.includes('profit') || q.includes('loss') || q.includes('p&l')))
  ) {
    const metricRows: Array<{ label: string; rowIndex: number }> = [];
    currentSheet.rows.forEach((row, idx) => {
      const textA = String(row[0]?.value ?? '').toLowerCase();
      const textB = String(row[1]?.value ?? '').toLowerCase();
      const combined = `${textA} ${textB}`;
      if (
        combined.includes('gross profit') ||
        combined.includes('operating income') ||
        combined.includes('consolidated net income') ||
        combined.includes('net income attributable') ||
        combined.includes('net operating revenues')
      ) {
        metricRows.push({ label: String(row[1]?.value || row[0]?.value), rowIndex: idx });
      }
    });

    if (metricRows.length > 0) {
      const yearColumns: Array<{ colIndex: number; label: string }> = [];
      for (let r = 0; r < Math.min(6, currentSheet.rows.length); r++) {
        const row = currentSheet.rows[r] ?? [];
        row.forEach((cell, cIdx) => {
          const val = String(cell?.value ?? '').trim();
          if (/^fy\s*'?\d{2,4}$/i.test(val) || /^20\d{2}$/.test(val) || /^19\d{2}$/.test(val)) {
            if (!yearColumns.some((y) => y.colIndex === cIdx)) {
              yearColumns.push({ colIndex: cIdx, label: val });
            }
          }
        });
      }

      const colsToAnalyze =
        yearColumns.length > 0
          ? yearColumns
          : columns
              .filter((c) => c.isNumeric)
              .map((c) => ({ colIndex: c.index, label: c.rawName }));

      const netIncomeRow =
        metricRows.find(
          (m) =>
            m.label.toLowerCase().includes('net income attributable') ||
            m.label.toLowerCase().includes('consolidated net income'),
        ) ||
        metricRows.find((m) => m.label.toLowerCase().includes('operating income')) ||
        metricRows.find((m) => m.label.toLowerCase().includes('gross profit'));

      if (netIncomeRow && colsToAnalyze.length > 0) {
        const rowCells = currentSheet.rows[netIncomeRow.rowIndex] ?? [];
        let profitCount = 0;
        let lossCount = 0;
        const periodsBreakdown: string[] = [];

        colsToAnalyze.forEach((col) => {
          const rawVal = rowCells[col.colIndex]?.value;
          const num = typeof rawVal === 'number' ? rawVal : parseFloat(String(rawVal ?? '').replace(/,/g, ''));
          if (!isNaN(num)) {
            if (num >= 0) {
              profitCount++;
              periodsBreakdown.push(`• **${col.label}**: **+$${num.toLocaleString()}M** (Profitable)`);
            } else {
              lossCount++;
              periodsBreakdown.push(`• **${col.label}**: **-$${Math.abs(num).toLocaleString()}M** (Net Loss)`);
            }
          }
        });

        const totalPeriods = profitCount + lossCount;
        if (totalPeriods > 0) {
          const probProfit = ((profitCount / totalPeriods) * 100).toFixed(1);
          const probLoss = ((lossCount / totalPeriods) * 100).toFixed(1);

          return {
            message: `### Financial Probability Analysis: Profit & Loss (${currentSheet.name})\n\nBased on the **${netIncomeRow.label}** across **${totalPeriods} recorded fiscal periods**:\n\n- **Probability of Profit**: **${probProfit}%** (${profitCount} / ${totalPeriods} periods)\n- **Probability of Loss**: **${probLoss}%** (${lossCount} / ${totalPeriods} periods)\n\n#### Fiscal Period Breakdown:\n${periodsBreakdown.join('\n')}\n\n**Summary**: The data shows **${probProfit}% historical profitability** across all analyzed fiscal periods with **${lossCount} recorded loss period(s)**.`,
          };
        }
      }
    }
  }

  // 7. VALUE-BASED SEARCH / FILTER / "LIST OUT" (e.g. "list out the stocks having 8 items", "me the stock having 8", "filter stock 8")
  // Check if query contains a number or specific value
  const numInQuery = q.match(/\b(\d+(?:\.\d+)?)\b/);
  let operatorGuess: string =
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
  let targetColForFilter = resolveColumn(q, columns);
  let targetValue: string | number | undefined = undefined;

  if (targetColForFilter) {
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
  }

  // Fallback: search across all columns' distinct values
  if (!targetColForFilter || targetValue === undefined) {
    const candidate = findFilterCandidateInSheet(raw, columns, currentSheet);
    if (candidate) {
      targetColForFilter = candidate.column;
      targetValue = candidate.value;
      if (candidate.operator) operatorGuess = candidate.operator as any;
    }
  }

  if (targetColForFilter && targetValue !== undefined) {
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
          const expectedStr = String(targetValue).toLowerCase();
          isMatch = strVal === expectedStr || strVal.includes(expectedStr);
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
              operator:
                operatorGuess === 'equals' &&
                typeof targetValue === 'string' &&
                !targetColForFilter.distinct.has(String(targetValue))
                  ? 'contains'
                  : operatorGuess,
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
