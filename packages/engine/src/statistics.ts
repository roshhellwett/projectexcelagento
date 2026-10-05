import { toNumericOrNull } from './formula/functions.js';

export interface NumericSummary {
  totalCount: number;
  numericCount: number;
  missingCount: number;
  nonNumericCount: number;
  sum: number | null;
  mean: number | null;
  min: number | null;
  max: number | null;
  median: number | null;
  q1: number | null;
  q3: number | null;
  sampleVariance: number | null;
  sampleStandardDeviation: number | null;
  populationStandardDeviation: number | null;
  iqr: number | null;
  lowerFence: number | null;
  upperFence: number | null;
  outlierCount: number;
  warnings: string[];
}

function finite(value: number): number | null {
  return Number.isFinite(value) ? value : null;
}

/** Linear-interpolated inclusive quantile, matching Excel PERCENTILE.INC (R type 7). */
function quantile(sorted: number[], probability: number): number | null {
  if (sorted.length === 0) return null;
  const position = (sorted.length - 1) * probability;
  const lower = Math.floor(position);
  const fraction = position - lower;
  const left = sorted[lower]!;
  const right = sorted[Math.min(lower + 1, sorted.length - 1)]!;
  return finite(left * (1 - fraction) + right * fraction);
}

/** Pure descriptive statistics with explicit exclusions and Tukey's 1.5-IQR outlier rule. */
export function summarizeNumericValues(values: Iterable<unknown>): NumericSummary {
  const numbers: number[] = [];
  let totalCount = 0;
  let missingCount = 0;
  let nonNumericCount = 0;
  let sum = 0;
  let compensation = 0;
  let mean = 0;
  let m2 = 0;
  for (const value of values) {
    totalCount += 1;
    if (value === null || value === undefined || (typeof value === 'string' && value.trim() === '')) {
      missingCount += 1;
      continue;
    }
    const number = toNumericOrNull(value);
    if (number === null) { nonNumericCount += 1; continue; }
    numbers.push(number);
    // Compensated summation and Welford variance avoid losing small contributions and
    // subtracting two almost equal, enormous squared totals.
    const corrected = number - compensation;
    const nextSum = sum + corrected;
    compensation = (nextSum - sum) - corrected;
    sum = nextSum;
    const delta = number - mean;
    mean += delta / numbers.length;
    m2 += delta * (number - mean);
  }
  numbers.sort((a, b) => a - b);
  const count = numbers.length;
  const q1 = quantile(numbers, 0.25);
  const q3 = quantile(numbers, 0.75);
  const iqr = q1 === null || q3 === null ? null : finite(q3 - q1);
  const lowerFence = q1 === null || iqr === null ? null : finite(q1 - 1.5 * iqr);
  const upperFence = q3 === null || iqr === null ? null : finite(q3 + 1.5 * iqr);
  const sampleVariance = count < 2 ? null : finite(Math.max(0, m2) / (count - 1));
  const result: NumericSummary = {
    totalCount, numericCount: count, missingCount, nonNumericCount,
    sum: count === 0 ? null : finite(sum),
    mean: count === 0 ? null : finite(mean),
    min: numbers[0] ?? null, max: numbers[count - 1] ?? null,
    median: quantile(numbers, 0.5), q1, q3, sampleVariance,
    sampleStandardDeviation: sampleVariance === null ? null : Math.sqrt(sampleVariance),
    populationStandardDeviation: count === 0 ? null : finite(Math.sqrt(Math.max(0, m2) / count)),
    iqr, lowerFence, upperFence,
    outlierCount: lowerFence === null || upperFence === null ? 0 : numbers.filter((number) => number < lowerFence || number > upperFence).length,
    warnings: [],
  };
  if (count === 0) result.warnings.push('No numeric observations; statistics are undefined.');
  else if (count === 1) result.warnings.push('Sample variance and standard deviation need at least two observations.');
  if (count > 0 && (result.sum === null || result.mean === null || (count > 1 && sampleVariance === null))) {
    result.warnings.push('Some statistics exceed the finite numeric range and are reported as undefined.');
  }
  return result;
}

export interface LinearRegressionResult {
  pairCount: number;
  excludedPairCount: number;
  slope: number | null;
  intercept: number | null;
  correlation: number | null;
  rSquared: number | null;
  residualStandardError: number | null;
  warnings: string[];
}

/** Ordinary least squares with an intercept. Only complete numeric pairs participate. */
export function fitLinearRegression(pairs: Iterable<readonly [unknown, unknown]>): LinearRegressionResult {
  let n = 0;
  let excluded = 0;
  let meanX = 0;
  let meanY = 0;
  let sxx = 0;
  let syy = 0;
  let sxy = 0;
  for (const [xValue, yValue] of pairs) {
    const x = toNumericOrNull(xValue);
    const y = toNumericOrNull(yValue);
    if (x === null || y === null) { excluded += 1; continue; }
    n += 1;
    const dx = x - meanX;
    const dy = y - meanY;
    meanX += dx / n;
    meanY += dy / n;
    sxx += dx * (x - meanX);
    syy += dy * (y - meanY);
    sxy += dx * (y - meanY);
  }
  const result: LinearRegressionResult = {
    pairCount: n, excludedPairCount: excluded, slope: null, intercept: null,
    correlation: null, rSquared: null, residualStandardError: null, warnings: [],
  };
  if (n < 2) { result.warnings.push('At least two complete numeric pairs are required.'); return result; }
  if (![meanX, meanY, sxx, syy, sxy].every(Number.isFinite)) {
    result.warnings.push('Values exceed the finite numeric range for regression.'); return result;
  }
  if (sxx <= 0) { result.warnings.push('The predictor is constant; slope and correlation are undefined.'); return result; }
  result.slope = finite(sxy / sxx);
  result.intercept = result.slope === null ? null : finite(meanY - result.slope * meanX);
  if (syy > 0) {
    result.correlation = Math.max(-1, Math.min(1, (sxy / Math.sqrt(sxx)) / Math.sqrt(syy)));
    result.rSquared = result.correlation ** 2;
  } else {
    result.warnings.push('The response is constant; correlation and R² are undefined.');
  }
  if (n > 2 && result.slope !== null) {
    result.residualStandardError = finite(Math.sqrt(Math.max(0, syy - result.slope * sxy) / (n - 2)));
  }
  return result;
}
