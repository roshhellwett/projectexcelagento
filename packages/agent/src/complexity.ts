const COMPLEX_SIGNALS = [
  'and then',
  'then ',
  'after that',
  'also ',
  'plus ',
  ', then',
  'first ',
  'secondly',
  'finally',
  'summarize',
  'summarise',
  'report',
  'insight',
  'compare',
  'trend',
  'breakdown',
  'analyze',
  'analyse',
  'transfer to',
  'migrate to',
  'move to',
  'populate',
  'extract to',
  'structure the sheet',
  'structure sheet',
  'reconcile',
  'reconciliation',
  'discrepancies',
  'flag discrepancies',
  'presentation-ready',
  'sales report',
  'multiple sheets',
  'in one sheet',
  'into one sheet',
  'one sheet',
  'single sheet',
  'combine sheets',
  'merge sheets',
  'consolidate',
  'all sheets into',
  'ek sheet',
  'sari sheet',
  'saari sheet',
  'sab ek sheet',
];

const DEPENDENT_VERB_COUNT =
  /\b(clean|merge|join|group|summarize|summarise|filter|sort|format|categorize|pivot|aggregate|count|sum|remove|add|rename|compute|transfer|migrate|populate|extract|structure|convert|reconcile|consolidate|combine|flatten|collapse)\b/gi;

/**
 * Decides whether a request warrants the full pipeline. Simple single-verb turns stay on the
 * fast path; a request that names several kinds of work or asks for analysis and action stays
 * under this gate so the planner does not spin up a team for "set cell A1 to 5".
 */
export function isComplexRequest(query: string): boolean {
  const verbs = [
    ...new Set(
      Array.from(query.matchAll(DEPENDENT_VERB_COUNT), (match) => match[0]!.toLowerCase()),
    ),
  ];
  if (verbs.length >= 2) return true;
  const lower = query.toLowerCase();
  if (COMPLEX_SIGNALS.some((signal) => lower.includes(signal))) return true;
  if (
    query.length > 140 &&
    /\b(clean|merge|join|group|pivot|aggregate|transfer|migrate|populate|extract|consolidate|combine)\b/i.test(query)
  )
    return true;
  return false;
}
