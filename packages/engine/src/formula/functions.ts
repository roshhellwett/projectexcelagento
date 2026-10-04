import type { FormulaValue } from './types.js';
import { isFormulaError, type FormulaErrorCode } from './errors.js';
import {
  coerceToDate,
  dateToExcelSerial,
  daysBetween,
  excelDate,
  formatIsoDate,
  parseUnambiguousDate,
} from './excel-date.js';

let formulaClock: () => Date = () => new Date();
let formulaDateSystem: '1900' | '1904' = '1900';

/** Overrides the wall clock used by TODAY()/NOW() so evaluation stays deterministic per call site. */
export function setFormulaClock(clock: () => Date): void {
  formulaClock = clock;
}

/**
 * Overrides which epoch date serials are counted from. Legacy Mac Excel workbooks use the
 * 1904 system, where the same serial denotes a date four years and one day earlier.
 */
export function setFormulaDateSystem(system: '1900' | '1904'): void {
  formulaDateSystem = system;
}

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
  if (isFormulaError(val)) return NaN;
  if (val instanceof Date) {
    const serial = dateToExcelSerial(val, formulaDateSystem);
    return serial ?? NaN;
  }
  if (typeof val === 'string') {
    const parsed = parseFloat(val.replace(/,/g, '').trim());
    return isNaN(parsed) ? 0 : parsed;
  }
  return 0;
}

/**
 * Resolves a date function's argument to a UTC `Date`, propagating an incoming error and
 * otherwise raising `#VALUE!`. This is the single gate every date function goes through, so
 * a raw Excel serial can no longer be mistaken for a calendar year.
 */
function toDateOrError(val: unknown): Date | FormulaErrorCode {
  if (isFormulaError(val)) return val;
  return coerceToDate(val, { dateSystem: formulaDateSystem }) ?? '#VALUE!';
}

/** Unwraps a `Date | FormulaErrorCode` into the date's components, or short-circuits on error. */
function dateParts(val: unknown): [number, number, number] | FormulaErrorCode {
  const date = toDateOrError(val);
  if (isFormulaError(date)) return date;
  return [date.getUTCFullYear(), date.getUTCMonth() + 1, date.getUTCDate()];
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
  if (val instanceof Date) return formatIsoDate(val);
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

/**
 * Weekdays between two dates, counted in whole weeks plus a bounded remainder walk.
 * The previous day-by-day loop took ~2.9M iterations for `NETWORKDAYS("1900-01-01","9999-12-31")`.
 */
function networkDaysBetween(start: Date, end: Date): number {
  const totalDays = daysBetween(start, end) + 1;
  if (totalDays <= 0) return 0;

  const startDow = start.getUTCDay();
  const wholeWeeks = Math.floor(totalDays / 7);
  let count = wholeWeeks * 5;

  for (let offset = wholeWeeks * 7; offset < totalDays; offset++) {
    const dow = (startDow + offset) % 7;
    if (dow !== 0 && dow !== 6) count++;
  }
  return count;
}

/**
 * Coerces a cell value to a number using the same rule the formula functions use, or null when
 * the value carries no magnitude. Booleans and dates are deliberately excluded: SUM, AVERAGE and
 * the COUNT family all ignore them, and an operation that disagreed with the formulas it sits
 * next to would quietly produce a different answer than `=SUM(A2:A10)`.
 */
export function toNumericOrNull(val: unknown): number | null {
  if (!isNumeric(val)) return null;
  return toNumber(val);
}

/** Evaluates criteria strings like ">10", "<=5", "<>Closed", "Active", or regex/wildcard */
export function matchesCriteria(val: unknown, criteria: unknown): boolean {
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
    const regexStr =
      '^' +
      critStr
        .replace(/[.+^${}()|[\]\\]/g, '\\$&')
        .replace(/\*/g, '.*')
        .replace(/\?/g, '.') +
      '$';
    return new RegExp(regexStr, 'i').test(toString(val));
  }

  if (isNumeric(critStr) && valNum !== null) {
    return valNum === parseFloat(critStr);
  }
  return toString(val).toLowerCase() === critStr.toLowerCase();
}

export type FormulaFunction = (...args: unknown[]) => FormulaValue;

/**
 * Excel propagates errors through aggregations: `SUM` over a range containing `#REF!`
 * returns `#REF!`, not a quietly reduced total. Ignoring errors is how a broken formula
 * becomes a plausible-looking wrong number.
 */
function firstErrorIn(values: unknown[]): FormulaErrorCode | null {
  for (const value of values) {
    if (isFormulaError(value)) return value;
  }
  return null;
}

/** Extremum without spreading the array, which overflows the stack on large ranges. */
function extremum(nums: number[], pick: (a: number, b: number) => number): number {
  if (nums.length === 0) return 0;
  return nums.reduce(pick);
}

export const FORMULA_FUNCTIONS: Record<string, FormulaFunction> = {
  // Math & Statistics
  SUM: (...args: unknown[]) => {
    const flat = flatten(args);
    const error = firstErrorIn(flat);
    if (error) return error;
    const nums = flat.filter(isNumeric).map(toNumber);
    return nums.reduce((a, b) => a + b, 0);
  },
  AVERAGE: (...args: unknown[]) => {
    const flat = flatten(args);
    const error = firstErrorIn(flat);
    if (error) return error;
    const nums = flat.filter(isNumeric).map(toNumber);
    return nums.length > 0 ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
  },
  MIN: (...args: unknown[]) => {
    const flat = flatten(args);
    const error = firstErrorIn(flat);
    if (error) return error;
    return extremum(flat.filter(isNumeric).map(toNumber), (a, b) => (a < b ? a : b));
  },
  MAX: (...args: unknown[]) => {
    const flat = flatten(args);
    const error = firstErrorIn(flat);
    if (error) return error;
    return extremum(flat.filter(isNumeric).map(toNumber), (a, b) => (a > b ? a : b));
  },
  COUNT: (...args: unknown[]) => {
    // COUNT counts numbers, so an error in the range is a value to skip, not a reason to fail.
    return flatten(args).filter(isNumeric).length;
  },
  COUNTA: (...args: unknown[]) => {
    return flatten(args).filter((v) => v !== null && v !== undefined && v !== '').length;
  },
  COUNTBLANK: (range: unknown) => {
    return flatten([range]).filter((v) => v === null || v === undefined || v === '').length;
  },
  MEDIAN: (...args: unknown[]) => {
    const flat = flatten(args);
    const error = firstErrorIn(flat);
    if (error) return error;
    const nums = flat
      .filter(isNumeric)
      .map(toNumber)
      .sort((a, b) => a - b);
    if (nums.length === 0) return 0;
    const mid = Math.floor(nums.length / 2);
    return nums.length % 2 !== 0 ? nums[mid]! : (nums[mid - 1]! + nums[mid]!) / 2;
  },
  ROUND: (num: unknown, digits: unknown = 0) => {
    const factor = Math.pow(10, toNumber(digits));
    return Math.round(toNumber(num) * factor) / factor;
  },
  ROUNDUP: (num: unknown, digits: unknown = 0) => {
    const factor = Math.pow(10, toNumber(digits));
    return (Math.ceil(Math.abs(toNumber(num)) * factor) / factor) * (toNumber(num) < 0 ? -1 : 1);
  },
  ROUNDDOWN: (num: unknown, digits: unknown = 0) => {
    const factor = Math.pow(10, toNumber(digits));
    return (Math.floor(Math.abs(toNumber(num)) * factor) / factor) * (toNumber(num) < 0 ? -1 : 1);
  },
  ABS: (num: unknown) => Math.abs(toNumber(num)),
  SQRT: (num: unknown) => {
    const n = toNumber(num);
    return n < 0 ? '#NUM!' : Math.sqrt(n);
  },
  POWER: (base: unknown, exp: unknown) => Math.pow(toNumber(base), toNumber(exp)),
  MOD: (n: unknown, d: unknown) => (toNumber(d) === 0 ? 0 : toNumber(n) % toNumber(d)),
  INT: (n: unknown) => Math.floor(toNumber(n)),
  TRUNC: (n: unknown, digits: unknown = 0) => {
    const factor = Math.pow(10, toNumber(digits));
    return Math.trunc(toNumber(n) * factor) / factor;
  },
  CEILING: (n: unknown, significance: unknown = 1) => {
    const sig = toNumber(significance);
    if (sig === 0) return 0;
    return Math.ceil(toNumber(n) / sig) * sig;
  },
  FLOOR: (n: unknown, significance: unknown = 1) => {
    const sig = toNumber(significance);
    if (sig === 0) return 0;
    return Math.floor(toNumber(n) / sig) * sig;
  },
  EXP: (n: unknown) => {
    const v = Math.exp(toNumber(n));
    return isFinite(v) ? v : '#NUM!';
  },
  LN: (n: unknown) => {
    const v = toNumber(n);
    return v <= 0 ? '#NUM!' : Math.log(v);
  },
  LOG: (n: unknown, base: unknown = 10) => {
    const v = toNumber(n);
    const b = toNumber(base);
    return v <= 0 || b <= 0 || b === 1 ? '#NUM!' : Math.log(v) / Math.log(b);
  },
  LOG10: (n: unknown) => {
    const v = toNumber(n);
    return v <= 0 ? '#NUM!' : Math.log10(v);
  },

  // Conditionals
  IF: (cond: unknown, trueVal: unknown, falseVal: unknown = false) => {
    return (toBoolean(cond) ? trueVal : falseVal) as FormulaValue;
  },
  IFS: (...args: unknown[]) => {
    for (let i = 0; i < args.length; i += 2) {
      if (toBoolean(args[i])) return (args[i + 1] ?? null) as FormulaValue;
    }
    return null;
  },
  IFERROR: (val: unknown, fallback: unknown) => {
    if (val === null || val === undefined) return fallback as FormulaValue;
    if (typeof val === 'number' && (isNaN(val) || !isFinite(val))) return fallback as FormulaValue;
    // Exact canonical-code match. A prefix test would swallow legitimate text such as the
    // SKU "#12345", which is exactly the kind of cell a user wraps in IFERROR.
    if (isFormulaError(val)) return fallback as FormulaValue;
    return val as FormulaValue;
  },
  AND: (...args: unknown[]) => {
    return flatten(args).every(toBoolean);
  },
  OR: (...args: unknown[]) => {
    return flatten(args).some(toBoolean);
  },
  NOT: (val: unknown) => !toBoolean(val),
  XOR: (...args: unknown[]) => {
    const count = flatten(args).filter(toBoolean).length;
    return count % 2 === 1;
  },

  // Conditional Aggregations
  SUMIF: (range: unknown, criteria: unknown, sumRange?: unknown) => {
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
  COUNTIF: (range: unknown, criteria: unknown) => {
    const flatRange = flatten([range]);
    return flatRange.filter((val) => matchesCriteria(val, criteria)).length;
  },
  AVERAGEIF: (range: unknown, criteria: unknown, avgRange?: unknown) => {
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
  SUMIFS: (sumRange: unknown, ...criteriaPairs: unknown[]) => {
    const flatSum = flatten([sumRange]);
    const pairs: { range: unknown[]; crit: unknown }[] = [];
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
  COUNTIFS: (...criteriaPairs: unknown[]) => {
    const pairs: { range: unknown[]; crit: unknown }[] = [];
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
  AVERAGEIFS: (avgRange: unknown, ...criteriaPairs: unknown[]) => {
    const flatAvg = flatten([avgRange]);
    const pairs: { range: unknown[]; crit: unknown }[] = [];
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
  VLOOKUP: (lookupVal: unknown, table: unknown, colIdx: unknown, exact: unknown = true) => {
    if (!Array.isArray(table) || table.length === 0) return '#N/A';
    const cIdx = toNumber(colIdx) - 1; // 1-indexed in Excel
    // Fourth argument mirrors Excel's range_lookup: FALSE (or omitted) = exact, TRUE = approximate
    const isExact = exact === undefined ? true : !toBoolean(exact);
    const targetStr = toString(lookupVal).trim().toLowerCase();

    if (isExact) {
      for (const row of table) {
        if (Array.isArray(row) && row.length > 0) {
          const key = toString(row[0]).trim().toLowerCase();
          if (key === targetStr) return (row[cIdx] ?? null) as FormulaValue;
        }
      }
      return '#N/A';
    }

    // Approximate match: largest key <= lookup value (table assumed sorted ascending)
    const numeric = (v: unknown): number | null => {
      if (typeof v === 'number' && isFinite(v)) return v;
      if (typeof v === 'string' && v.trim() !== '' && isNumeric(v)) return Number(v);
      return null;
    };
    const targetNum = numeric(lookupVal);
    let best: unknown = '#N/A';
    let bestFound = false;
    for (const row of table) {
      if (Array.isArray(row) && row.length > 0) {
        const keyRaw = row[0];
        const keyNum = numeric(keyRaw);
        const leq =
          targetNum !== null && keyNum !== null
            ? keyNum <= targetNum
            : toString(keyRaw).trim().toLowerCase() <= targetStr;
        if (leq) {
          best = (row[cIdx] ?? null) as FormulaValue;
          bestFound = true;
        } else {
          break;
        }
      }
    }
    return bestFound ? (best as FormulaValue) : '#N/A';
  },
  XLOOKUP: (
    lookupVal: unknown,
    lookupArray: unknown,
    returnArray: unknown,
    notFound: unknown = '#N/A',
  ) => {
    const flatLookup = flatten([lookupArray]);
    const flatReturn = flatten([returnArray]);
    const target = toString(lookupVal).trim().toLowerCase();

    for (let i = 0; i < flatLookup.length; i++) {
      if (toString(flatLookup[i]).trim().toLowerCase() === target) {
        return (flatReturn[i] ?? null) as FormulaValue;
      }
    }
    return notFound as FormulaValue;
  },
  INDEX: (array: unknown, rowNum: unknown, colNum: unknown = 1) => {
    const r = toNumber(rowNum) - 1;
    const c = toNumber(colNum) - 1;
    if (Array.isArray(array)) {
      if (Array.isArray(array[0])) {
        return (array[r]?.[c] ?? '#REF!') as FormulaValue;
      }
      return (array[r] ?? '#REF!') as FormulaValue;
    }
    return array as FormulaValue;
  },
  MATCH: (lookupVal: unknown, lookupArray: unknown, matchType: unknown = 0) => {
    const flat = flatten([lookupArray]);
    const target = toString(lookupVal).trim().toLowerCase();
    const type = toNumber(matchType);

    if (type === 0) {
      for (let i = 0; i < flat.length; i++) {
        if (toString(flat[i]).trim().toLowerCase() === target) return i + 1;
      }
      return '#N/A';
    }

    const numericVal = (v: unknown): number | null => {
      if (typeof v === 'number' && isFinite(v)) return v;
      if (typeof v === 'string' && v.trim() !== '' && isNumeric(v)) return Number(v);
      return null;
    };
    const targetNum = numericVal(lookupVal);
    let best = -1;
    for (let i = 0; i < flat.length; i++) {
      const keyNum = numericVal(flat[i]);
      const leq =
        targetNum !== null && keyNum !== null
          ? keyNum <= targetNum
          : toString(flat[i]).trim().toLowerCase() <= target;
      const geq =
        targetNum !== null && keyNum !== null
          ? keyNum >= targetNum
          : toString(flat[i]).trim().toLowerCase() >= target;
      if (type === 1 && leq) best = i;
      if (type === -1 && geq && best === -1) best = i;
    }
    return best >= 0 ? best + 1 : '#N/A';
  },
  CHOOSE: (index: unknown, ...choices: unknown[]) => {
    const idx = toNumber(index) - 1;
    return (choices[idx] ?? '#VALUE!') as FormulaValue;
  },
  IFNA: (val: unknown, fallback: unknown) =>
    val === '#N/A' || (typeof val === 'string' && val.toUpperCase() === '#N/A')
      ? (fallback as FormulaValue)
      : (val as FormulaValue),
  SWITCH: (expr: unknown, ...pairs: unknown[]) => {
    for (let i = 0; i + 1 < pairs.length; i += 2) {
      if (toString(expr) === toString(pairs[i]) || toNumber(expr) === toNumber(pairs[i])) {
        return pairs[i + 1] as FormulaValue;
      }
    }
    return pairs.length % 2 === 1 ? (pairs[pairs.length - 1] as FormulaValue) : '#N/A';
  },
  SUMPRODUCT: (...arrays: unknown[]) => {
    const flats = arrays.map((a) => flatten([a]));
    if (flats.length === 0 || flats.some((f) => f.length !== flats[0]!.length)) return '#VALUE!';
    let sum = 0;
    for (let i = 0; i < flats[0]!.length; i++) {
      let product = 1;
      for (const f of flats) product *= toNumber(f[i]);
      sum += product;
    }
    return sum;
  },
  GEOMEAN: (...args: unknown[]) => {
    const nums = flatten(args)
      .map(toNumber)
      .filter((n) => n > 0);
    return nums.length === 0
      ? '#NUM!'
      : Math.pow(
          nums.reduce((a, b) => a * b, 1),
          1 / nums.length,
        );
  },
  STDEV: (...args: unknown[]) => {
    const nums = flatten(args).flatMap((v) => (isNumeric(v) ? [toNumber(v)] : []));
    if (nums.length < 2) return '#DIV/0!';
    const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
    return Math.sqrt(nums.reduce((a, b) => a + (b - mean) ** 2, 0) / (nums.length - 1));
  },
  VAR: (...args: unknown[]) => {
    const nums = flatten(args).flatMap((v) => (isNumeric(v) ? [toNumber(v)] : []));
    if (nums.length < 2) return '#DIV/0!';
    const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
    return nums.reduce((a, b) => a + (b - mean) ** 2, 0) / (nums.length - 1);
  },
  MODE: (...args: unknown[]) => {
    const nums = flatten(args).flatMap((v) => (isNumeric(v) ? [toNumber(v)] : []));
    if (nums.length === 0) return '#N/A';
    const counts = new Map<number, number>();
    for (const n of nums) counts.set(n, (counts.get(n) ?? 0) + 1);
    let best: number | null = null;
    let bestCount = 1;
    for (const [n, c] of counts) {
      if (c > bestCount) {
        best = n;
        bestCount = c;
      }
    }
    return best ?? '#N/A';
  },
  LARGE: (data: unknown, k: unknown) => {
    const nums = flatten([data])
      .flatMap((v) => (isNumeric(v) ? [toNumber(v)] : []))
      .sort((a, b) => b - a);
    const idx = toNumber(k) - 1;
    return nums[idx] ?? '#NUM!';
  },
  SMALL: (data: unknown, k: unknown) => {
    const nums = flatten([data])
      .flatMap((v) => (isNumeric(v) ? [toNumber(v)] : []))
      .sort((a, b) => a - b);
    const idx = toNumber(k) - 1;
    return nums[idx] ?? '#NUM!';
  },
  RANK: (value: unknown, data: unknown, order: unknown = 0) => {
    const nums = flatten([data]).flatMap((v) => (isNumeric(v) ? [toNumber(v)] : []));
    const target = toNumber(value);
    return order
      ? nums.filter((n) => n < target).length + 1
      : nums.filter((n) => n > target).length + 1;
  },
  PERCENTILE: (data: unknown, k: unknown) => {
    const nums = flatten([data])
      .flatMap((v) => (isNumeric(v) ? [toNumber(v)] : []))
      .sort((a, b) => a - b);
    if (nums.length === 0) return '#NUM!';
    const p = toNumber(k);
    if (p < 0 || p > 1) return '#NUM!';
    const rank = p * (nums.length - 1);
    const lo = Math.floor(rank);
    const hi = Math.ceil(rank);
    return lo === hi ? nums[lo]! : nums[lo]! + (nums[hi]! - nums[lo]!) * (rank - lo);
  },
  FV: (rate: unknown, nper: unknown, pmt: unknown, pv: unknown = 0, type: unknown = 0) => {
    const r = toNumber(rate);
    const n = toNumber(nper);
    const p = toNumber(pmt);
    const present = toNumber(pv);
    if (r === 0) return -(present + p * n);
    const factor = Math.pow(1 + r, n);
    return -(present * factor + p * (1 + r * toNumber(type)) * ((factor - 1) / r));
  },
  PV: (rate: unknown, nper: unknown, pmt: unknown, fv: unknown = 0, type: unknown = 0) => {
    const r = toNumber(rate);
    const n = toNumber(nper);
    const p = toNumber(pmt);
    const future = toNumber(fv);
    if (r === 0) return -(future + p * n);
    const factor = Math.pow(1 + r, n);
    return -(future + p * (1 + r * toNumber(type)) * ((factor - 1) / r)) / factor;
  },
  PMT: (rate: unknown, nper: unknown, pv: unknown, fv: unknown = 0, type: unknown = 0) => {
    const r = toNumber(rate);
    const n = toNumber(nper);
    const present = toNumber(pv);
    const future = toNumber(fv);
    if (n === 0) return '#DIV/0!';
    if (r === 0) return -(present + future) / n;
    const factor = Math.pow(1 + r, n);
    return (-(present * factor + future) * r) / ((1 + r * toNumber(type)) * (factor - 1));
  },
  NPER: (rate: unknown, pmt: unknown, pv: unknown, fv: unknown = 0, type: unknown = 0) => {
    const r = toNumber(rate);
    const p = toNumber(pmt);
    const present = toNumber(pv);
    const future = toNumber(fv);
    if (r === 0) return p === 0 ? '#DIV/0!' : -(present + future) / p;
    const adjustedP = p * (1 + r * toNumber(type));
    return Math.log((adjustedP - future * r) / (present * r + adjustedP)) / Math.log(1 + r);
  },
  NPV: (rate: unknown, ...values: unknown[]) => {
    const r = toNumber(rate);
    const flows = flatten(values).map(toNumber);
    return flows.reduce((sum, cf, i) => sum + cf / Math.pow(1 + r, i + 1), 0);
  },
  IRR: (values: unknown, guess: unknown = 0.1) => {
    const flows = flatten([values]).map(toNumber);
    let rate = toNumber(guess);
    for (let iter = 0; iter < 100; iter++) {
      let npv = 0;
      let dnpv = 0;
      for (let i = 0; i < flows.length; i++) {
        npv += flows[i]! / Math.pow(1 + rate, i);
        dnpv -= (i * flows[i]!) / Math.pow(1 + rate, i + 1);
      }
      if (Math.abs(dnpv) < 1e-12) return '#NUM!';
      const next = rate - npv / dnpv;
      if (Math.abs(next - rate) < 1e-10) return next;
      rate = next;
    }
    return '#NUM!';
  },
  RATE: (
    nper: unknown,
    pmt: unknown,
    pv: unknown,
    fv: unknown = 0,
    type: unknown = 0,
    guess: unknown = 0.1,
  ) => {
    let rate = toNumber(guess);
    const n = toNumber(nper);
    const p = toNumber(pmt);
    const present = toNumber(pv);
    const future = toNumber(fv);
    for (let iter = 0; iter < 100; iter++) {
      const factor = Math.pow(1 + rate, n);
      const npv =
        present * factor + p * (1 + rate * toNumber(type)) * ((factor - 1) / rate) + future;
      const dFactor = n * Math.pow(1 + rate, n - 1);
      const dNpv =
        present * dFactor +
        p * toNumber(type) * ((factor - 1) / rate) +
        p * (1 + rate * toNumber(type)) * ((dFactor * rate - (factor - 1)) / (rate * rate));
      if (Math.abs(dNpv) < 1e-12) return '#NUM!';
      const next = rate - npv / dNpv;
      if (Math.abs(next - rate) < 1e-10) return next;
      rate = next;
    }
    return '#NUM!';
  },
  DATEDIF: (start: unknown, end: unknown, unit: unknown) => {
    const a = toDateOrError(start);
    if (isFormulaError(a)) return a;
    const b = toDateOrError(end);
    if (isFormulaError(b)) return b;
    const u = toString(unit).toUpperCase();
    if (u === 'D') return daysBetween(b, a);
    if (u === 'M')
      return (
        (b.getUTCFullYear() - a.getUTCFullYear()) * 12 +
        (b.getUTCMonth() - a.getUTCMonth()) -
        (b.getUTCDate() < a.getUTCDate() ? 1 : 0)
      );
    if (u === 'Y')
      return (
        b.getUTCFullYear() -
        a.getUTCFullYear() -
        (b.getUTCMonth() < a.getUTCMonth() ||
        (b.getUTCMonth() === a.getUTCMonth() && b.getUTCDate() < a.getUTCDate())
          ? 1
          : 0)
      );
    return '#NUM!';
  },
  DAYS: (end: unknown, start: unknown) => {
    const a = toDateOrError(end);
    if (isFormulaError(a)) return a;
    const b = toDateOrError(start);
    return isFormulaError(b) ? b : daysBetween(a, b);
  },
  EOMONTH: (start: unknown, months: unknown) => {
    const d = toDateOrError(start);
    if (isFormulaError(d)) return d;
    // Day 0 of the following month is the last day of the target month.
    const end = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + toNumber(months) + 1, 0));
    return Number.isFinite(end.getTime()) ? end : '#NUM!';
  },
  EDATE: (start: unknown, months: unknown) => {
    const d = toDateOrError(start);
    if (isFormulaError(d)) return d;
    const shifted = new Date(
      Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + toNumber(months), d.getUTCDate()),
    );
    return Number.isFinite(shifted.getTime()) ? shifted : '#NUM!';
  },
  WEEKDAY: (dateVal: unknown, type: unknown = 1) => {
    const d = toDateOrError(dateVal);
    if (isFormulaError(d)) return d;
    const dow = d.getUTCDay();
    const mode = toNumber(type);
    // Type 3 is Excel's Monday-first numbering; the default is Sunday-first 1..7.
    return mode === 3 ? (dow === 0 ? 6 : dow - 1) : dow + 1;
  },
  NETWORKDAYS: (start: unknown, end: unknown) => {
    const a = toDateOrError(start);
    if (isFormulaError(a)) return a;
    const b = toDateOrError(end);
    if (isFormulaError(b)) return b;
    // Iterating one day at a time across a multi-century span would hang the tab, so this
    // counts whole weeks and only walks the remainder.
    if (daysBetween(a, b) < 0) return Math.abs(networkDaysBetween(b, a));
    return networkDaysBetween(a, b);
  },
  HOUR: (dateVal: unknown) => {
    const d = toDateOrError(dateVal);
    return isFormulaError(d) ? d : d.getUTCHours();
  },
  MINUTE: (dateVal: unknown) => {
    const d = toDateOrError(dateVal);
    return isFormulaError(d) ? d : d.getUTCMinutes();
  },
  SECOND: (dateVal: unknown) => {
    const d = toDateOrError(dateVal);
    return isFormulaError(d) ? d : d.getUTCSeconds();
  },
  HLOOKUP: (lookupVal: unknown, table: unknown, rowIdx: unknown, exact: unknown = true) => {
    if (!Array.isArray(table) || table.length === 0) return '#N/A';
    const rIdx = toNumber(rowIdx) - 1;
    const isExact = exact === undefined ? true : !toBoolean(exact);
    const target = toString(lookupVal).trim().toLowerCase();
    const header = table[0] as unknown[];
    for (let c = 0; c < header.length; c++) {
      if (toString(header[c]).trim().toLowerCase() === target) {
        return ((table[rIdx] as unknown[])?.[c] ?? '#REF!') as FormulaValue;
      }
    }
    if (!isExact) {
      let best = -1;
      for (let c = 0; c < header.length; c++) {
        if (toString(header[c]).trim().toLowerCase() <= target) best = c;
      }
      if (best >= 0) return ((table[rIdx] as unknown[])?.[best] ?? '#REF!') as FormulaValue;
    }
    return '#N/A';
  },
  CHAR: (n: unknown) => String.fromCharCode(toNumber(n)),
  CODE: (text: unknown) => toString(text).charCodeAt(0) || 0,
  REPT: (text: unknown, n: unknown) => toString(text).repeat(Math.max(0, toNumber(n))),
  CLEAN: (text: unknown) =>
    // eslint-disable-next-line no-control-regex
    toString(text).replace(/[\x00-\x1F\x7F]/g, ''),

  // Text
  CONCAT: (...args: unknown[]) => {
    return flatten(args).map(toString).join('');
  },
  CONCATENATE: (...args: unknown[]) => {
    return flatten(args).map(toString).join('');
  },
  TEXTJOIN: (delimiter: unknown, ignoreEmpty: unknown, ...args: unknown[]) => {
    const sep = toString(delimiter);
    const skip = toBoolean(ignoreEmpty);
    const flat = flatten(args).map(toString);
    const filtered = skip ? flat.filter((s) => s.trim().length > 0) : flat;
    return filtered.join(sep);
  },
  LEFT: (text: unknown, numChars: unknown = 1) => {
    return toString(text).slice(0, Math.max(0, toNumber(numChars)));
  },
  RIGHT: (text: unknown, numChars: unknown = 1) => {
    const str = toString(text);
    const n = Math.max(0, toNumber(numChars));
    return str.slice(Math.max(0, str.length - n));
  },
  MID: (text: unknown, start: unknown, length: unknown) => {
    const str = toString(text);
    const s = Math.max(1, toNumber(start)) - 1;
    const len = Math.max(0, toNumber(length));
    return str.slice(s, s + len);
  },
  LEN: (text: unknown) => toString(text).length,
  UPPER: (text: unknown) => toString(text).toUpperCase(),
  LOWER: (text: unknown) => toString(text).toLowerCase(),
  PROPER: (text: unknown) => {
    return toString(text).replace(/\b\w/g, (c) => c.toUpperCase());
  },
  TRIM: (text: unknown) => toString(text).trim().replace(/\s+/g, ' '),
  SUBSTITUTE: (text: unknown, oldText: unknown, newText: unknown, instanceNum?: unknown) => {
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
  REPLACE: (oldText: unknown, startNum: unknown, numChars: unknown, newText: unknown) => {
    const str = toString(oldText);
    const start = Math.max(1, toNumber(startNum)) - 1;
    const len = Math.max(0, toNumber(numChars));
    return str.slice(0, start) + toString(newText) + str.slice(start + len);
  },
  FIND: (findText: unknown, withinText: unknown, startNum: unknown = 1) => {
    const start = Math.max(1, toNumber(startNum)) - 1;
    const idx = toString(withinText).indexOf(toString(findText), start);
    return idx >= 0 ? idx + 1 : '#VALUE!';
  },
  SEARCH: (findText: unknown, withinText: unknown, startNum: unknown = 1) => {
    const start = Math.max(1, toNumber(startNum)) - 1;
    const idx = toString(withinText).toLowerCase().indexOf(toString(findText).toLowerCase(), start);
    return idx >= 0 ? idx + 1 : '#VALUE!';
  },
  EXACT: (t1: unknown, t2: unknown) => toString(t1) === toString(t2),
  TEXT: (val: unknown, fmt?: unknown) => {
    const format = fmt === undefined || fmt === null ? '' : toString(fmt);
    // A bare number is only treated as a date serial when the format actually asks for
    // date parts, so TEXT(1234, "0.00") still formats as a number.
    const looksDate = /[yMdhs]/.test(format.replace(/"[^"]*"/g, ''));
    const dateVal = looksDate ? coerceToDate(val, { dateSystem: formulaDateSystem }) : null;
    if (dateVal && !isNaN(dateVal.getTime())) {
      if (!format) return formatIsoDate(dateVal);
      const pad = (n: number) => String(n).padStart(2, '0');
      return format
        .replace(/YYYY/g, String(dateVal.getUTCFullYear()))
        .replace(/YY/g, String(dateVal.getUTCFullYear()).slice(-2))
        .replace(/MMMM/g, dateVal.toLocaleString('en-US', { month: 'long', timeZone: 'UTC' }))
        .replace(/MMM/g, dateVal.toLocaleString('en-US', { month: 'short', timeZone: 'UTC' }))
        .replace(/MM/g, pad(dateVal.getUTCMonth() + 1))
        .replace(/DD/g, pad(dateVal.getUTCDate()))
        .replace(/HH/g, pad(dateVal.getUTCHours()))
        .replace(/mm/g, pad(dateVal.getUTCMinutes()))
        .replace(/ss/g, pad(dateVal.getUTCSeconds()));
    }
    const n = toNumber(val);
    if (typeof val === 'number' || (typeof val === 'string' && isNumeric(val))) {
      if (format.includes('%')) {
        const decimals = format.split('.')[1]?.replace(/[^0#]/g, '').length ?? 0;
        return (n * 100).toFixed(decimals) + '%';
      }
      const decimals = format.split('.')[1]?.replace(/[^0#]/g, '').length;
      if (decimals !== undefined) return n.toFixed(decimals);
    }
    return toString(val);
  },
  VALUE: (text: unknown) => toNumber(text),

  // Date & Time
  TODAY: () => {
    // UTC components throughout: mixing local getters with a UTC-built Date would make the
    // result depend on the viewer's timezone and shift the day for anyone east of Greenwich.
    const now = formulaClock();
    return new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate()));
  },
  NOW: () => formulaClock(),
  DATE: (y: unknown, m: unknown, d: unknown) => {
    const date = excelDate(toNumber(y), toNumber(m), toNumber(d));
    return date ?? '#VALUE!';
  },
  DATEVALUE: (val: unknown) => {
    const parsed = parseUnambiguousDate(toString(val));
    return parsed ?? '#VALUE!';
  },
  YEAR: (dateVal: unknown) => {
    const parts = dateParts(dateVal);
    return isFormulaError(parts) ? parts : parts[0];
  },
  MONTH: (dateVal: unknown) => {
    const parts = dateParts(dateVal);
    return isFormulaError(parts) ? parts : parts[1];
  },
  DAY: (dateVal: unknown) => {
    const parts = dateParts(dateVal);
    return isFormulaError(parts) ? parts : parts[2];
  },

  // Information. Without this family a defensive formula cannot test what it is holding,
  // which is why so much real-world spreadsheet logic reaches for ISERROR/ISNUMBER.
  ISBLANK: (val: unknown) => val === null || val === undefined || val === '',
  ISNUMBER: (val: unknown) => typeof val === 'number' && !isNaN(val),
  ISTEXT: (val: unknown) => typeof val === 'string' && !isFormulaError(val),
  ISNONTEXT: (val: unknown) => !(typeof val === 'string' && !isFormulaError(val)),
  ISLOGICAL: (val: unknown) => typeof val === 'boolean',
  ISERROR: (val: unknown) => isFormulaError(val),
  ISERR: (val: unknown) => isFormulaError(val) && val !== '#N/A',
  ISNA: (val: unknown) => val === '#N/A',
  NA: () => '#N/A',
  N: (val: unknown) => {
    if (isFormulaError(val)) return val;
    if (val instanceof Date) return dateToExcelSerial(val, formulaDateSystem) ?? 0;
    return toNumber(val);
  },
  T: (val: unknown) => (typeof val === 'string' && !isFormulaError(val) ? val : ''),
  TYPE: (val: unknown) => {
    if (isFormulaError(val)) return 16;
    if (val instanceof Date) return 16;
    if (typeof val === 'boolean') return 4;
    if (typeof val === 'number' || val === null || val === undefined) return 1;
    return 2;
  },
};
