import { useId, useMemo } from 'react';
import type { Workbook } from '@excel-agent/engine';
import {
  compareWorkbookBaseline,
  createAnalystBriefing,
  type BaselineCellSnapshot,
  type BriefingChart,
  type WorkbookBaselineComparison,
} from '@excel-agent/agent';
import '../briefing.css';

export interface AnalystBriefingProps {
  workbook: Workbook;
  sheetName: string;
  /** An explicitly supplied snapshot; nothing is persisted or inferred by this component. */
  baseline?: Workbook;
}

const numberFormat = new Intl.NumberFormat('en-US', { maximumSignificantDigits: 8 });
const formatNumber = (value: number | null) =>
  value === null ? 'Undefined' : numberFormat.format(value);
const delta = (value: number) => (value > 0 ? `+${value}` : String(value));

/** Standalone, provider-free briefing. All content derives from the supplied workbook snapshots. */
export function AnalystBriefing({ workbook, sheetName, baseline }: AnalystBriefingProps) {
  const briefing = useMemo(() => createAnalystBriefing(workbook, sheetName), [workbook, sheetName]);
  const comparison = useMemo(
    () => (baseline ? compareWorkbookBaseline(baseline, workbook) : null),
    [baseline, workbook],
  );
  const quality = briefing.quality;
  return (
    <section className="analyst-briefing" aria-label={`Analyst briefing for ${briefing.sheetName}`}>
      <header className="briefing-hero">
        <div>
          <span className="briefing-eyebrow">DETERMINISTIC ANALYSIS · READ ONLY</span>
          <h2>Analyst briefing</h2>
          <p>Your workbook, with the evidence attached. Descriptive data—not a forecast.</p>
        </div>
        <span className="briefing-badge">Local computation · No model required</span>
      </header>

      <div className="briefing-source-strip">
        <div>
          <span>WORKBOOK SOURCE</span>
          <strong>{briefing.sheetName}</strong>
        </div>
        <div>
          <span>FULL STORED RANGE</span>
          <code>{briefing.sourceRange?.a1 ?? 'No stored range'}</code>
        </div>
        <div>
          <span>DATA RANGE · ROW 1 ASSUMED HEADER</span>
          <code>{briefing.dataRange?.a1 ?? 'No data rows'}</code>
        </div>
        <div>
          <span>DATE EPOCH</span>
          <strong>{briefing.dateSystem}</strong>
        </div>
      </div>

      {briefing.status === 'sheet-not-found' ? (
        <p className="briefing-empty" role="status">
          Sheet “{sheetName}” was not found. Choose an existing sheet to inspect.
        </p>
      ) : (
        <>
          <div className="briefing-metrics">
            <Metric
              label="Data rows"
              value={formatNumber(briefing.dataRowCount)}
              detail="All stored rows after the assumed header"
            />
            <Metric
              label="Nonblank completeness"
              value={
                quality.completenessPercent === null
                  ? 'Undefined'
                  : `${formatNumber(quality.completenessPercent)}%`
              }
              detail="Presence only—not correctness"
            />
            <Metric
              label="Missing cells"
              value={formatNumber(quality.missingCount)}
              detail={`Of ${formatNumber(quality.totalCells)} data cells, including ragged rows`}
            />
            <Metric
              label="Numeric exclusions"
              value={formatNumber(quality.nonNumericCount)}
              detail="Nonnumeric cells; missing cells counted separately"
            />
          </div>

          <section className="briefing-card" aria-label="Data quality summary">
            <div className="briefing-section-heading">
              <div>
                <span className="briefing-eyebrow">01 · DATA QUALITY</span>
                <h3>What deserves attention</h3>
              </div>
              <span className="briefing-badge">Full-range audit</span>
            </div>
            <ul className="briefing-findings">
              {briefing.findings.map((finding) => (
                <li key={finding}>{finding}</li>
              ))}
            </ul>
            <dl className="briefing-quality-counts">
              <div>
                <dt>Numeric observations</dt>
                <dd>{quality.numericCount}</dd>
              </div>
              <div>
                <dt>Numeric text included</dt>
                <dd>{quality.numericTextCount}</dd>
              </div>
              <div>
                <dt>Formula cells</dt>
                <dd>{quality.formulaCount}</dd>
              </div>
              <div>
                <dt>Error values excluded</dt>
                <dd>{quality.formulaErrorCount}</dd>
              </div>
              <div>
                <dt>Unsupported formulas / dependents</dt>
                <dd>{quality.unsupportedFormulaCount}</dd>
              </div>
              <div>
                <dt>Clock/random formulas / dependents</dt>
                <dd>{quality.volatileFormulaCount}</dd>
              </div>
              <div>
                <dt>Ragged data rows</dt>
                <dd>{quality.raggedRowCount}</dd>
              </div>
              <div>
                <dt>Fully blank data rows</dt>
                <dd>{quality.fullyBlankRowCount}</dd>
              </div>
              <div>
                <dt>Blank headers</dt>
                <dd>{quality.blankHeaderCount}</dd>
              </div>
              <div>
                <dt>Repeated header labels</dt>
                <dd>
                  {quality.duplicateHeaders.length ? quality.duplicateHeaders.join(', ') : 'None'}
                </dd>
              </div>
            </dl>
            {quality.examples.length > 0 && (
              <details className="briefing-details">
                <summary>
                  Source evidence · {quality.examples.length} quality examples
                  {quality.omittedExampleCount > 0
                    ? ` · ${quality.omittedExampleCount} more omitted`
                    : ''}
                </summary>
                <ul className="briefing-evidence-list">
                  {quality.examples.map((example) => (
                    <li key={example.address}>
                      <code>{example.address}</code>
                      <span>{example.reason}</span>
                    </li>
                  ))}
                </ul>
              </details>
            )}
          </section>

          <section className="briefing-card" aria-label="Numeric distributions">
            <div className="briefing-section-heading">
              <div>
                <span className="briefing-eyebrow">02 · NUMERIC DISTRIBUTIONS</span>
                <h3>Every number has a source</h3>
              </div>
              <span className="briefing-badge">{briefing.columnCount} columns</span>
            </div>
            <p className="briefing-description">
              Independent column totals, not combined metrics. Dates and booleans are excluded.
              Numeric text is included; missing cells and all other exclusions are counted
              separately.
            </p>
            {briefing.columns.length ? (
              <div
                className="briefing-table-wrap"
                tabIndex={0}
                role="region"
                aria-label="Scrollable full-column statistics"
              >
                <table className="briefing-table briefing-numeric-table">
                  <caption>
                    Full-column descriptive statistics · data rows only · displayed numbers rounded
                    to 8 significant digits
                  </caption>
                  <thead>
                    <tr>
                      {[
                        'Column / source',
                        'Total cells',
                        'Numeric',
                        'Missing',
                        'Nonnumeric excluded',
                        'Sum',
                        'Mean',
                        'Median',
                        'Min',
                        'Q1',
                        'Q3',
                        'Max',
                        'Sample SD',
                        'IQR outliers',
                      ].map((name) => (
                        <th key={name} scope="col">
                          {name}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {briefing.columns.map((column) => (
                      <tr key={column.column}>
                        <th scope="row">
                          <strong>
                            {column.header} <span>({column.column})</span>
                          </strong>
                          <code>{column.sourceRange?.a1 ?? 'No data range'}</code>
                          <small>Header: {column.headerSource.a1}</small>
                        </th>
                        {[
                          column.numeric.totalCount,
                          column.numeric.numericCount,
                          column.numeric.missingCount,
                          column.numeric.nonNumericCount,
                          column.numeric.sum,
                          column.numeric.mean,
                          column.numeric.median,
                          column.numeric.min,
                          column.numeric.q1,
                          column.numeric.q3,
                          column.numeric.max,
                          column.numeric.sampleStandardDeviation,
                          column.numeric.outlierCount,
                        ].map((value, index) => (
                          <td
                            key={index}
                            title={value === null ? 'No defined statistic' : String(value)}
                          >
                            {formatNumber(value)}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <p className="briefing-empty">No columns to summarize.</p>
            )}
            {briefing.columns.some((column) => column.numeric.warnings.length) && (
              <ul className="briefing-warnings">
                {briefing.columns.flatMap((column) =>
                  column.numeric.warnings.map((warning) => (
                    <li key={`${column.column}-${warning}`}>
                      {column.header} ({column.column}): {warning}
                    </li>
                  )),
                )}
              </ul>
            )}
            <p className="briefing-footnote">
              Quartiles use inclusive interpolation. Outliers use Tukey’s 1.5 × IQR rule; a flagged
              observation is not proof of an error. Sample SD needs at least two numeric
              observations.
            </p>
          </section>

          <section aria-label="Distribution chart previews">
            <div className="briefing-section-heading">
              <div>
                <span className="briefing-eyebrow">03 · VISUAL EVIDENCE</span>
                <h3>Distribution, not prediction</h3>
              </div>
              <span className="briefing-badge">Bounded previews</span>
            </div>
            <p className="briefing-description">
              Counts come from the full source ranges. Each chart includes its exact count table;
              omitted categories are not silently grouped or estimated.
            </p>
            <div className="briefing-chart-grid">
              {briefing.charts.map((chart) => (
                <DistributionChart key={chart.id} chart={chart} />
              ))}
            </div>
            {briefing.charts.length === 0 && (
              <p className="briefing-empty">
                No chartable numeric or categorical observations below the assumed header.
              </p>
            )}
            {briefing.omittedChartCount > 0 && (
              <p className="briefing-footnote">
                {briefing.omittedChartCount} additional chart previews omitted. The statistics table
                still includes every column.
              </p>
            )}
          </section>
        </>
      )}

      {comparison ? (
        <BaselineComparison comparison={comparison} />
      ) : (
        <section className="briefing-card briefing-baseline-empty" aria-label="Baseline comparison">
          <h3>Baseline comparison</h3>
          <p>
            No baseline supplied. This is a snapshot, not a trend. Supply an earlier workbook
            snapshot for exact address-based changes.
          </p>
        </section>
      )}

      <section className="briefing-methods" aria-label="Methods and limits">
        <h3>Methods & honest limits</h3>
        <ul>
          {briefing.limits.map((limit) => (
            <li key={limit}>{limit}</li>
          ))}
        </ul>
      </section>
    </section>
  );
}

function Metric({ label, value, detail }: { label: string; value: string; detail: string }) {
  return (
    <div className="briefing-metric">
      <span>{label}</span>
      <strong>{value}</strong>
      <p>{detail}</p>
    </div>
  );
}

function DistributionChart({ chart }: { chart: BriefingChart }) {
  const id = useId();
  const max = Math.max(1, ...chart.series.map((item) => item.count));
  const slot = 320 / Math.max(1, chart.series.length);
  return (
    <article className="briefing-card briefing-chart-card">
      <div className="briefing-chart-heading">
        <h4>
          {chart.header} <span>({chart.column})</span>
        </h4>
        <span className="briefing-eyebrow">
          {chart.kind === 'histogram' ? 'NUMERIC FREQUENCY' : 'TOP CATEGORIES'}
        </span>
      </div>
      <p className="briefing-chart-source">
        Source: <code>{chart.sourceRange.a1}</code>
      </p>
      <svg
        viewBox="0 0 360 170"
        className="briefing-chart"
        role="img"
        aria-labelledby={`${id}-title`}
        aria-describedby={`${id}-desc`}
        focusable="false"
      >
        <title id={`${id}-title`}>
          {chart.header} ({chart.column}){' '}
          {chart.kind === 'histogram' ? 'numeric distribution' : 'categorical counts'}
        </title>
        <desc id={`${id}-desc`}>
          Bar heights show counts, not time or predictions. {chart.includedCount} observations
          shown; {chart.omittedCount} categorical observations omitted. Exact labels and counts are
          in the table below. Source: {chart.sourceRange.a1}.
        </desc>
        <line x1="20" y1="140" x2="340" y2="140" className="briefing-chart-axis" />
        {chart.series.map((item, index) => {
          const height = (item.count / max) * 108;
          return (
            <g key={index}>
              <rect
                x={20 + index * slot + 5}
                y={140 - height}
                width={Math.max(2, slot - 10)}
                height={height}
                rx="3"
                className="briefing-chart-bar"
              >
                <title>
                  {item.label}: {item.count}
                </title>
              </rect>
              <text
                x={20 + index * slot + slot / 2}
                y={130 - height}
                textAnchor="middle"
                className="briefing-chart-count"
              >
                {item.count}
              </text>
              <text
                x={20 + index * slot + slot / 2}
                y="158"
                textAnchor="middle"
                className="briefing-chart-tick"
              >
                {index + 1}
              </text>
            </g>
          );
        })}
      </svg>
      <div className="briefing-table-wrap">
        <table className="briefing-table briefing-chart-table">
          <caption>
            {chart.header} ({chart.column}) chart data · exact counts
          </caption>
          <thead>
            <tr>
              <th scope="col">Bar</th>
              <th scope="col">
                {chart.kind === 'histogram' ? 'Numeric bin' : 'Category (value type)'}
              </th>
              <th scope="col">Count</th>
            </tr>
          </thead>
          <tbody>
            {chart.series.map((item, index) => (
              <tr key={index}>
                <td>{index + 1}</td>
                <th scope="row">{item.label}</th>
                <td>{item.count}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <p className="briefing-footnote">
        {chart.includedCount} observations shown · {chart.omittedCount} omitted.
      </p>
      {chart.notes.map((note) => (
        <p key={note} className="briefing-footnote">
          {note}
        </p>
      ))}
    </article>
  );
}

function describeSnapshot(snapshot: BaselineCellSnapshot) {
  if (!snapshot.present) return 'Absent cell';
  return JSON.stringify(
    snapshot,
    (_key, value: unknown) =>
      typeof value === 'number' && Object.is(value, -0) ? '-0 (number)' : value,
    2,
  );
}

function BaselineComparison({ comparison }: { comparison: WorkbookBaselineComparison }) {
  return (
    <section className="briefing-card" aria-label="Baseline comparison">
      <div className="briefing-section-heading">
        <div>
          <span className="briefing-eyebrow">04 · SNAPSHOT COMPARISON</span>
          <h3>What changed, exactly</h3>
        </div>
        <span className="briefing-badge">Address-based evidence</span>
      </div>
      <p className="briefing-description">
        {comparison.changedCellCount} stored cells changed. Rows include headers. Date epoch:{' '}
        {comparison.baselineDateSystem} → {comparison.currentDateSystem}
        {comparison.dateSystemChanged ? ' (changed)' : ' (unchanged)'}.
      </p>
      <div
        className="briefing-table-wrap"
        tabIndex={0}
        role="region"
        aria-label="Scrollable sheet comparison"
      >
        <table className="briefing-table">
          <caption>Sheet presence and shape deltas · not comparable metric changes</caption>
          <thead>
            <tr>
              {[
                'Sheet',
                'Presence',
                'Rows before → after',
                'Row delta',
                'Columns before → after',
                'Column delta',
                'Layout',
              ].map((label) => (
                <th key={label} scope="col">
                  {label}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {comparison.sheets.map((sheet) => (
              <tr key={sheet.sheetName}>
                <th scope="row">
                  {sheet.sheetName}
                  <small>
                    Baseline: {sheet.baselineSourceRange?.a1 ?? 'Absent / no stored range'}
                  </small>
                  <small>
                    Current: {sheet.currentSourceRange?.a1 ?? 'Absent / no stored range'}
                  </small>
                </th>
                <td>{sheet.presence}</td>
                <td>
                  {sheet.baselineRows} → {sheet.currentRows}
                </td>
                <td>{delta(sheet.rowDelta)}</td>
                <td>
                  {sheet.baselineColumns} → {sheet.currentColumns}
                </td>
                <td>{delta(sheet.columnDelta)}</td>
                <td>
                  {sheet.layoutChanged ? 'Changed—positional only' : 'Same shape & headers'}
                  {sheet.headerChanged && <small>Header changed</small>}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <ul className="briefing-warnings">
        {comparison.notes.map((note) => (
          <li key={note}>{note}</li>
        ))}
      </ul>
      {comparison.changedCells.length > 0 && (
        <details className="briefing-details" open>
          <summary>
            Exact changed-cell evidence · {comparison.changedCells.length} examples ·{' '}
            {comparison.omittedChangedCellCount} additional changes omitted
          </summary>
          <div
            className="briefing-table-wrap"
            tabIndex={0}
            role="region"
            aria-label="Scrollable changed-cell evidence"
          >
            <table className="briefing-table briefing-change-table">
              <caption>
                Exact stored-cell examples · formula values are caches, not recalculated analytics
              </caption>
              <thead>
                <tr>
                  <th scope="col">Address</th>
                  <th scope="col">Baseline cell</th>
                  <th scope="col">Current cell</th>
                </tr>
              </thead>
              <tbody>
                {comparison.changedCells.map((change) => (
                  <tr key={change.address}>
                    <th scope="row">
                      <code>{change.address}</code>
                    </th>
                    <td>
                      <pre>{describeSnapshot(change.before)}</pre>
                    </td>
                    <td>
                      <pre>{describeSnapshot(change.after)}</pre>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </details>
      )}
      {comparison.changedCellCount === 0 && (
        <p className="briefing-empty">
          No exact stored-cell changes. Epoch metadata is reported separately; this does not assert
          unchanged evaluated formulas or business performance.
        </p>
      )}
    </section>
  );
}
