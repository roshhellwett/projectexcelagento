import type { FormulaValue } from './types.js';

function flatten(args: unknown[]): unknown[] {
  const result: unknown[] = [];
  for (const item of args) {
    if (Array.isArray(item)) {
      result.push(...flatten(item));
    } else {
      result.push(item);
    }
  }
  return result;
}

function toNumber(val: unknown): number {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  if (typeof val === 'boolean') return val ? 1 : 0;
  if (val instanceof Date) return val.getTime();
  if (typeof val === 'string') {
    const parsed = parseFloat(val.replace(/,/g, '').trim());
    return isNaN(parsed) ? 0 : parsed;
  }
  return 0;
}

function isNumeric(val: unknown): boolean {
  if (typeof val === 'number') return !isNaN(val);
  if (typeof val === 'string') {
    return val.trim() !== '' && !isNaN(Number(val.replace(/,/g, '')));
  }
  return false;
}

function toString(val: unknown): string {
  if (val === null || val === undefined) return '';
  if (val instanceof Date) return val.toISOString().slice(0, 10);
  return String(val);
}

function toBoolean(val: unknown): boolean {
  if (typeof val === 'boolean') return val;
  if (typeof val === 'number') return val !== 0;
  if (typeof val === 'string') {
    const s = val.trim().toUpperCase();
    if (s === 'TRUE') return true;
    if (s === 'FALSE') return false;
    return s.length > 0;
  }
  return Boolean(val);
}

/** Evaluates criteria strings like ">10", "<=5", "<>Closed", "Active", or regex/wildcard */
function matchesCriteria(val: unknown, criteria: unknown): boolean {
  const critStr = toString(criteria).trim();
  const valNum = isNumeric(val) ? toNumber(val) : null;

  if (critStr.startsWith('>=')) {
    const target = parseFloat(critStr.slice(2));
    return valNum !== null && valNum >= target;
  }
  if (critStr.startsWith('<=')) {
    const target = parseFloat(critStr.slice(2));
    return valNum !== null && valNum <= target;
  }
  if (critStr.startsWith('<>')) {
    const target = critStr.slice(2).trim();
    if (isNumeric(target) && valNum !== null) return valNum !== parseFloat(target);
    return toString(val).toLowerCase() !== target.toLowerCase();
  }
  if (critStr.startsWith('>')) {
    const target = parseFloat(critStr.slice(1));
    return valNum !== null && valNum > target;
  }
  if (critStr.startsWith('<')) {
    const target = parseFloat(critStr.slice(1));
    return valNum !== null && valNum < target;
  }
  if (critStr.startsWith('=')) {
    const target = critStr.slice(1).trim();
    if (isNumeric(target) && valNum !== null) return valNum === parseFloat(target);
    return toString(val).toLowerCase() === target.toLowerCase();
  }

  // Exact or wildcard comparison
  if (critStr.includes('*') || critStr.includes('?')) {
    const regexStr = '^' + critStr.replace(/\*/g, '.*').replace(/\?/g, '.') + '$';
    return new RegExp(regexStr, 'i').test(toString(val));
  }

  if (isNumeric(critStr) && valNum !== null) {
    return valNum === parseFloat(critStr);
  }
  return toString(val).toLowerCase() === critStr.toLowerCase();
}

export type FormulaFunction = (...args: any[]) => FormulaValue;

export const FORMULA_FUNCTIONS: Record<string, FormulaFunction> = {
  // Math & Statistics
  SUM: (...args: any[]) => {
    const nums = flatten(args).filter(isNumeric).map(toNumber);
    return nums.reduce((a, b) => a + b, 0);
  },
  AVERAGE: (...args: any[]) => {
    const nums = flatten(args).filter(isNumeric).map(toNumber);
    return nums.length > 0 ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
  },
  MIN: (...args: any[]) => {
    const nums = flatten(args).filter(isNumeric).map(toNumber);
    return nums.length > 0 ? Math.min(...nums) : 0;
  },
  MAX: (...args: any[]) => {
    const nums = flatten(args).filter(isNumeric).map(toNumber);
    return nums.length > 0 ? Math.max(...nums) : 0;
  },
  COUNT: (...args: any[]) => {
    return flatten(args).filter(isNumeric).length;
  },
  COUNTA: (...args: any[]) => {
    return flatten(args).filter((v) => v !== null && v !== undefined && v !== '').length;
  },
  COUNTBLANK: (range: any) => {
    return flatten([range]).filter((v) => v === null || v === undefined || v === '').length;
  },
  MEDIAN: (...args: any[]) => {
    const nums = flatten(args).filter(isNumeric).map(toNumber).sort((a, b) => a - b);
    if (nums.length === 0) return 0;
    const mid = Math.floor(nums.length / 2);
    return nums.length % 2 !== 0 ? nums[mid]! : (nums[mid - 1]! + nums[mid]!) / 2;
  },
  ROUND: (num: any, digits: any = 0) => {
    const factor = Math.pow(10, toNumber(digits));
    return Math.round(toNumber(num) * factor) / factor;
  },
  ROUNDUP: (num: any, digits: any = 0) => {
    const factor = Math.pow(10, toNumber(digits));
    return (Math.ceil(Math.abs(toNumber(num)) * factor) / factor) * (toNumber(num) < 0 ? -1 : 1);
  },
  ROUNDDOWN: (num: any, digits: any = 0) => {
    const factor = Math.pow(10, toNumber(digits));
    return (Math.floor(Math.abs(toNumber(num)) * factor) / factor) * (toNumber(num) < 0 ? -1 : 1);
  },
  ABS: (num: any) => Math.abs(toNumber(num)),
  SQRT: (num: any) => Math.sqrt(Math.max(0, toNumber(num))),
  POWER: (base: any, exp: any) => Math.pow(toNumber(base), toNumber(exp)),
  MOD: (n: any, d: any) => (toNumber(d) === 0 ? 0 : toNumber(n) % toNumber(d)),
  INT: (n: any) => Math.floor(toNumber(n)),
  TRUNC: (n: any, digits: any = 0) => {
    const factor = Math.pow(10, toNumber(digits));
    return Math.trunc(toNumber(n) * factor) / factor;
  },
  CEILING: (n: any, significance: any = 1) => {
    const sig = toNumber(significance);
    if (sig === 0) return 0;
    return Math.ceil(toNumber(n) / sig) * sig;
  },
  FLOOR: (n: any, significance: any = 1) => {
    const sig = toNumber(significance);
    if (sig === 0) return 0;
    return Math.floor(toNumber(n) / sig) * sig;
  },
  EXP: (n: any) => Math.exp(toNumber(n)),
  LN: (n: any) => Math.log(toNumber(n)),
  LOG: (n: any, base: any = 10) => Math.log(toNumber(n)) / Math.log(toNumber(base)),
  LOG10: (n: any) => Math.log10(toNumber(n)),

  // Conditionals
  IF: (cond: any, trueVal: any, falseVal: any = false) => {
    return toBoolean(cond) ? trueVal : falseVal;
  },
  IFS: (...args: any[]) => {
    for (let i = 0; i < args.length; i += 2) {
      if (toBoolean(args[i])) return args[i + 1] ?? null;
    }
    return null;
  },
  IFERROR: (val: any, fallback: any) => {
    if (val === null || val === undefined) return fallback;
    if (typeof val === 'number' && (isNaN(val) || !isFinite(val))) return fallback;
    if (typeof val === 'string' && val.startsWith('#')) return fallback;
    return val;
  },
  AND: (...args: any[]) => {
    return flatten(args).every(toBoolean);
  },
  OR: (...args: any[]) => {
    return flatten(args).some(toBoolean);
  },
  NOT: (val: any) => !toBoolean(val),
  XOR: (...args: any[]) => {
    const count = flatten(args).filter(toBoolean).length;
    return count % 2 === 1;
  },

  // Conditional Aggregations
  SUMIF: (range: any[], criteria: any, sumRange?: any[]) => {
    const flatRange = flatten([range]);
    const flatSum = sumRange ? flatten([sumRange]) : flatRange;
    let sum = 0;
    for (let i = 0; i < flatRange.length; i++) {
      if (matchesCriteria(flatRange[i], criteria)) {
        sum += toNumber(flatSum[i]);
      }
    }
    return sum;
  },
  COUNTIF: (range: any[], criteria: any) => {
    const flatRange = flatten([range]);
    return flatRange.filter((val) => matchesCriteria(val, criteria)).length;
  },
  AVERAGEIF: (range: any[], criteria: any, avgRange?: any[]) => {
    const flatRange = flatten([range]);
    const flatAvg = avgRange ? flatten([avgRange]) : flatRange;
    let sum = 0;
    let count = 0;
    for (let i = 0; i < flatRange.length; i++) {
      if (matchesCriteria(flatRange[i], criteria)) {
        sum += toNumber(flatAvg[i]);
        count++;
      }
    }
    return count > 0 ? sum / count : 0;
  },
  SUMIFS: (sumRange: any[], ...criteriaPairs: any[]) => {
    const flatSum = flatten([sumRange]);
    const pairs: { range: any[]; crit: any }[] = [];
    for (let i = 0; i < criteriaPairs.length; i += 2) {
      pairs.push({ range: flatten([criteriaPairs[i]]), crit: criteriaPairs[i + 1] });
    }
    let sum = 0;
    for (let i = 0; i < flatSum.length; i++) {
      const match = pairs.every((p) => matchesCriteria(p.range[i], p.crit));
      if (match) sum += toNumber(flatSum[i]);
    }
    return sum;
  },
  COUNTIFS: (...criteriaPairs: any[]) => {
    const pairs: { range: any[]; crit: any }[] = [];
    for (let i = 0; i < criteriaPairs.length; i += 2) {
      pairs.push({ range: flatten([criteriaPairs[i]]), crit: criteriaPairs[i + 1] });
    }
    if (pairs.length === 0) return 0;
    const len = pairs[0]!.range.length;
    let count = 0;
    for (let i = 0; i < len; i++) {
      if (pairs.every((p) => matchesCriteria(p.range[i], p.crit))) {
        count++;
      }
    }
    return count;
  },
  AVERAGEIFS: (avgRange: any[], ...criteriaPairs: any[]) => {
    const flatAvg = flatten([avgRange]);
    const pairs: { range: any[]; crit: any }[] = [];
    for (let i = 0; i < criteriaPairs.length; i += 2) {
      pairs.push({ range: flatten([criteriaPairs[i]]), crit: criteriaPairs[i + 1] });
    }
    let sum = 0;
    let count = 0;
    for (let i = 0; i < flatAvg.length; i++) {
      if (pairs.every((p) => matchesCriteria(p.range[i], p.crit))) {
        sum += toNumber(flatAvg[i]);
        count++;
      }
    }
    return count > 0 ? sum / count : 0;
  },

  // Lookup & Reference
  VLOOKUP: (lookupVal: any, table: any, colIdx: any, exact: any = true) => {
    if (!Array.isArray(table) || table.length === 0) return '#N/A';
    const cIdx = toNumber(colIdx) - 1; // 1-indexed in Excel
    const isExact = exact === undefined || toBoolean(exact);
    const targetStr = toString(lookupVal).trim().toLowerCase();

    for (const row of table) {
      if (Array.isArray(row) && row.length > 0) {
        const key = toString(row[0]).trim().toLowerCase();
        if (isExact ? key === targetStr : key.includes(targetStr)) {
          return row[cIdx] ?? null;
        }
      }
    }
    return '#N/A';
  },
  XLOOKUP: (lookupVal: any, lookupArray: any[], returnArray: any[], notFound: any = '#N/A') => {
    const flatLookup = flatten([lookupArray]);
    const flatReturn = flatten([returnArray]);
    const target = toString(lookupVal).trim().toLowerCase();

    for (let i = 0; i < flatLookup.length; i++) {
      if (toString(flatLookup[i]).trim().toLowerCase() === target) {
        return flatReturn[i] ?? null;
      }
    }
    return notFound;
  },
  INDEX: (array: any, rowNum: any, colNum: any = 1) => {
    const r = toNumber(rowNum) - 1;
    const c = toNumber(colNum) - 1;
    if (Array.isArray(array)) {
      if (Array.isArray(array[0])) {
        return array[r]?.[c] ?? '#REF!';
      }
      return array[r] ?? '#REF!';
    }
    return array;
  },
  MATCH: (lookupVal: any, lookupArray: any[], _matchType: any = 0) => {
    const flat = flatten([lookupArray]);
    const target = toString(lookupVal).trim().toLowerCase();
    for (let i = 0; i < flat.length; i++) {
      if (toString(flat[i]).trim().toLowerCase() === target) {
        return i + 1; // 1-indexed in Excel
      }
    }
    return '#N/A';
  },
  CHOOSE: (index: any, ...choices: any[]) => {
    const idx = toNumber(index) - 1;
    return choices[idx] ?? '#VALUE!';
  },

  // Text
  CONCAT: (...args: any[]) => {
    return flatten(args).map(toString).join('');
  },
  CONCATENATE: (...args: any[]) => {
    return flatten(args).map(toString).join('');
  },
  TEXTJOIN: (delimiter: any, ignoreEmpty: any, ...args: any[]) => {
    const sep = toString(delimiter);
    const skip = toBoolean(ignoreEmpty);
    const flat = flatten(args).map(toString);
    const filtered = skip ? flat.filter((s) => s.trim().length > 0) : flat;
    return filtered.join(sep);
  },
  LEFT: (text: any, numChars: any = 1) => {
    return toString(text).slice(0, Math.max(0, toNumber(numChars)));
  },
  RIGHT: (text: any, numChars: any = 1) => {
    const str = toString(text);
    const n = Math.max(0, toNumber(numChars));
    return str.slice(Math.max(0, str.length - n));
  },
  MID: (text: any, start: any, length: any) => {
    const str = toString(text);
    const s = Math.max(1, toNumber(start)) - 1;
    const len = Math.max(0, toNumber(length));
    return str.slice(s, s + len);
  },
  LEN: (text: any) => toString(text).length,
  UPPER: (text: any) => toString(text).toUpperCase(),
  LOWER: (text: any) => toString(text).toLowerCase(),
  PROPER: (text: any) => {
    return toString(text).replace(/\b\w/g, (c) => c.toUpperCase());
  },
  TRIM: (text: any) => toString(text).trim().replace(/\s+/g, ' '),
  SUBSTITUTE: (text: any, oldText: any, newText: any, instanceNum?: any) => {
    const str = toString(text);
    const target = toString(oldText);
    const rep = toString(newText);
    if (!target) return str;
    if (instanceNum !== undefined) {
      let count = 0;
      const targetNum = toNumber(instanceNum);
      return str.replace(new RegExp(target.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), (m) => {
        count++;
        return count === targetNum ? rep : m;
      });
    }
    return str.replaceAll(target, rep);
  },
  REPLACE: (oldText: any, startNum: any, numChars: any, newText: any) => {
    const str = toString(oldText);
    const start = Math.max(1, toNumber(startNum)) - 1;
    const len = Math.max(0, toNumber(numChars));
    return str.slice(0, start) + toString(newText) + str.slice(start + len);
  },
  FIND: (findText: any, withinText: any, startNum: any = 1) => {
    const start = Math.max(1, toNumber(startNum)) - 1;
    const idx = toString(withinText).indexOf(toString(findText), start);
    return idx >= 0 ? idx + 1 : '#VALUE!';
  },
  SEARCH: (findText: any, withinText: any, startNum: any = 1) => {
    const start = Math.max(1, toNumber(startNum)) - 1;
    const idx = toString(withinText).toLowerCase().indexOf(toString(findText).toLowerCase(), start);
    return idx >= 0 ? idx + 1 : '#VALUE!';
  },
  EXACT: (t1: any, t2: any) => toString(t1) === toString(t2),
  TEXT: (val: any, _fmt?: any) => {
    if (val instanceof Date) return val.toISOString().slice(0, 10);
    return toString(val);
  },
  VALUE: (text: any) => toNumber(text),

  // Date & Time
  TODAY: () => new Date().toISOString().slice(0, 10),
  NOW: () => new Date().toISOString(),
  DATE: (y: any, m: any, d: any) => {
    const date = new Date(Date.UTC(toNumber(y), toNumber(m) - 1, toNumber(d)));
    return date.toISOString().slice(0, 10);
  },
  YEAR: (dateVal: any) => {
    const d = new Date(toString(dateVal));
    return isNaN(d.getTime()) ? 0 : d.getUTCFullYear();
  },
  MONTH: (dateVal: any) => {
    const d = new Date(toString(dateVal));
    return isNaN(d.getTime()) ? 0 : d.getUTCMonth() + 1;
  },
  DAY: (dateVal: any) => {
    const d = new Date(toString(dateVal));
    return isNaN(d.getTime()) ? 0 : d.getUTCDate();
  },
};
