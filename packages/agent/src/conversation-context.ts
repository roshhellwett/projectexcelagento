import type { ChatMessage } from './types.js';

/** The only message roles that are safe and useful in conversational context. */
export type ConversationContextRole = Extract<ChatMessage['role'], 'user' | 'assistant'>;

/** Where a message in a context window came from. */
export type ConversationContextSource = 'recent' | 'recalled' | 'current-query';

/** Trust marker for consumers that want to apply their own prompt boundaries. */
export type ConversationContextTrust = 'recent' | 'lower-trust' | 'current';

/** A provider-compatible message with provenance for later orchestrator integration. */
export interface ConversationContextMessage {
  role: ConversationContextRole;
  content: string;
  source: ConversationContextSource;
  trust: ConversationContextTrust;
  /** The source index in the supplied history. The current query has no history index. */
  originalIndex?: number;
  /** Set only for recalled messages, in deterministic ranking order. */
  overlapTokens?: string[];
  /** Set only for recalled messages. Higher scores indicate stronger lexical relevance. */
  relevanceScore?: number;
}

/** A recalled older user turn, including the evidence used to select it. */
export interface RecalledConversationMessage extends ConversationContextMessage {
  source: 'recalled';
  trust: 'lower-trust';
  overlapTokens: string[];
  relevanceScore: number;
}

/** Limits applied while constructing a conversation context window. */
export interface ConversationContextOptions {
  /** Maximum number of output messages, including the current query when present. */
  maxMessages?: number;
  /** Number of eligible user/assistant messages retained from the end of history. */
  recentMessageCount?: number;
  /** Maximum number of older user messages considered for lexical recall. */
  maxRecalledMessages?: number;
  /** Maximum number of characters across all output message contents. */
  maxCharacters?: number;
  /** Maximum characters in any single output message. */
  maxMessageCharacters?: number;
  /** Minimum number of exact query-token matches required for an older turn. */
  minOverlapTokens?: number;
}

/** The bounded result consumed by a future orchestrator/provider adapter. */
export interface ConversationContext {
  /** Chronological, provider-compatible messages; the current query is last. */
  messages: ConversationContextMessage[];
  /** Recalled messages in relevance order, after all output caps are applied. */
  recalled: RecalledConversationMessage[];
  /** Number of eligible history messages left out of the bounded window. */
  omittedMessageCount: number;
  /** Sum of message-content lengths in `messages`. */
  characterCount: number;
}

const DEFAULTS = {
  maxMessages: 16,
  recentMessageCount: 8,
  maxRecalledMessages: 4,
  maxCharacters: 12_000,
  maxMessageCharacters: 2_000,
  minOverlapTokens: 1,
} as const;

// These are deliberately separate from the configurable defaults. Even a caller that passes a
// very large number cannot turn this helper into an unbounded prompt accumulator.
const HARD_LIMITS = {
  maxMessages: 32,
  maxRecalledMessages: 16,
  maxCharacters: 24_000,
  maxMessageCharacters: 4_000,
} as const;

const STOP_WORDS = new Set([
  'a',
  'an',
  'and',
  'are',
  'as',
  'at',
  'be',
  'been',
  'but',
  'by',
  'can',
  'could',
  'for',
  'from',
  'how',
  'i',
  'if',
  'in',
  'is',
  'it',
  'me',
  'my',
  'of',
  'on',
  'or',
  'our',
  'please',
  'that',
  'the',
  'their',
  'then',
  'these',
  'they',
  'this',
  'those',
  'to',
  'was',
  'we',
  'were',
  'what',
  'when',
  'where',
  'which',
  'who',
  'why',
  'with',
  'would',
  'you',
  'your',
]);

const TOKEN_PATTERN = /[\p{L}\p{N}]+(?:['’][\p{L}\p{N}]+)?/gu;

const CORRECTION_PATTERN =
  /\b(?:actually|correction|to clarify|that(?:'|’)?s wrong|that is wrong|instead|rather|undo|revert|do not|don't|never|stop using|change (?:this|that) to|ignore my previous)\b/i;
const PREFERENCE_PATTERN =
  /\b(?:prefer|preference|always|from now on|by default|keep|usually|i want|i need)\b/i;

const LOWER_TRUST_PREFIX =
  '[LOWER-TRUST CONVERSATIONAL CONTEXT: prior user message; background only]';
const LOWER_TRUST_SUFFIX = '[END LOWER-TRUST CONVERSATIONAL CONTEXT]';

interface EligibleMessage {
  role: ConversationContextRole;
  content: string;
  originalIndex: number;
}

interface RankedRecall extends RecalledConversationMessage {
  correctionPriority: number;
  overlapCount: number;
  overlapRatio: number;
}

interface NormalizedOptions {
  maxMessages: number;
  recentMessageCount: number;
  maxRecalledMessages: number;
  maxCharacters: number;
  maxMessageCharacters: number;
  minOverlapTokens: number;
}

function boundedInteger(value: number | undefined, fallback: number, maximum: number): number {
  if (value === undefined || !Number.isFinite(value)) return fallback;
  return Math.max(0, Math.min(maximum, Math.floor(value)));
}

function normalizeOptions(options: ConversationContextOptions): NormalizedOptions {
  return {
    maxMessages: Math.max(
      1,
      boundedInteger(options.maxMessages, DEFAULTS.maxMessages, HARD_LIMITS.maxMessages),
    ),
    recentMessageCount: boundedInteger(
      options.recentMessageCount,
      DEFAULTS.recentMessageCount,
      HARD_LIMITS.maxMessages,
    ),
    maxRecalledMessages: boundedInteger(
      options.maxRecalledMessages,
      DEFAULTS.maxRecalledMessages,
      HARD_LIMITS.maxRecalledMessages,
    ),
    maxCharacters: Math.max(
      1,
      boundedInteger(options.maxCharacters, DEFAULTS.maxCharacters, HARD_LIMITS.maxCharacters),
    ),
    maxMessageCharacters: Math.max(
      1,
      boundedInteger(
        options.maxMessageCharacters,
        DEFAULTS.maxMessageCharacters,
        HARD_LIMITS.maxMessageCharacters,
      ),
    ),
    minOverlapTokens: Math.max(
      1,
      boundedInteger(options.minOverlapTokens, DEFAULTS.minOverlapTokens, 64),
    ),
  };
}

/**
 * Removes role-like control characters and delimiter lookalikes from conversation text.
 *
 * History is user-provided input. Keeping it in one physical line prevents a past message from
 * manufacturing a new textual role turn, while marker escaping prevents it from closing the
 * lower-trust wrapper used for recalled material.
 */
function sanitizeConversationText(value: string): string {
  const withoutControls = Array.from(value.normalize('NFKC'), (character) => {
    const codePoint = character.codePointAt(0) ?? 0;
    return codePoint <= 0x1f ||
      (codePoint >= 0x7f && codePoint <= 0x9f) ||
      (codePoint >= 0x200b && codePoint <= 0x200d) ||
      codePoint === 0xfeff
      ? ' '
      : character;
  }).join('');

  return withoutControls
    .replace(/\s+/g, ' ')
    .replace(/\[LOWER-TRUST CONVERSATIONAL CONTEXT\]/gi, '[escaped lower-trust marker]')
    .replace(/\[END LOWER-TRUST CONVERSATIONAL CONTEXT\]/gi, '[escaped end marker]')
    .trim();
}

function truncateText(value: string, maxCharacters: number): string {
  if (value.length <= maxCharacters) return value;
  if (maxCharacters <= 1) return value.slice(0, maxCharacters);
  return `${value.slice(0, maxCharacters - 1)}…`;
}

function conversationTokens(value: string): Set<string> {
  const normalized = value.normalize('NFKC').toLowerCase().replace(/’/g, "'");
  const matches = normalized.match(TOKEN_PATTERN) ?? [];
  return new Set(matches.filter((token) => token.length > 1 && !STOP_WORDS.has(token)));
}

function correctionPriority(value: string): number {
  if (CORRECTION_PATTERN.test(value)) return 2;
  if (PREFERENCE_PATTERN.test(value)) return 1;
  return 0;
}

function wrapRecalledContent(
  content: string,
  role: ConversationContextRole,
  maxCharacters: number,
): string {
  const label = `${LOWER_TRUST_PREFIX} Role: ${role}.`;
  const availableForBody = maxCharacters - label.length - LOWER_TRUST_SUFFIX.length - 2;
  if (availableForBody <= 0) return truncateText(label, maxCharacters);

  const body = truncateText(content, availableForBody);
  return `${label} ${body} ${LOWER_TRUST_SUFFIX}`;
}

function toRecentMessage(
  message: EligibleMessage,
  maxCharacters: number,
): ConversationContextMessage {
  return {
    role: message.role,
    content: truncateText(message.content, maxCharacters),
    source: 'recent',
    trust: 'recent',
    originalIndex: message.originalIndex,
  };
}

function rankOlderUserMessage(
  message: EligibleMessage,
  queryTokens: Set<string>,
  minOverlapTokens: number,
): RankedRecall | undefined {
  if (message.role !== 'user') return undefined;

  const matchedTokens = [...conversationTokens(message.content)].filter((token) =>
    queryTokens.has(token),
  );
  if (matchedTokens.length < minOverlapTokens) return undefined;

  const priority = correctionPriority(message.content);
  const overlapRatio = matchedTokens.length / Math.max(queryTokens.size, 1);
  return {
    role: message.role,
    content: message.content,
    source: 'recalled',
    trust: 'lower-trust',
    originalIndex: message.originalIndex,
    overlapTokens: matchedTokens,
    // The explicit correction/preference class intentionally outranks raw overlap count. A
    // later correction should not be displaced by a longer, stale instruction with more words.
    relevanceScore: priority * 100 + matchedTokens.length * 10 + overlapRatio,
    correctionPriority: priority,
    overlapCount: matchedTokens.length,
    overlapRatio,
  };
}

function compareRecalls(left: RankedRecall, right: RankedRecall): number {
  if (left.correctionPriority !== right.correctionPriority)
    return right.correctionPriority - left.correctionPriority;
  if (left.overlapCount !== right.overlapCount) return right.overlapCount - left.overlapCount;
  if (left.overlapRatio !== right.overlapRatio) return right.overlapRatio - left.overlapRatio;
  return (right.originalIndex ?? -1) - (left.originalIndex ?? -1);
}

function fitMessageToBudget(
  message: ConversationContextMessage,
  remainingCharacters: number,
): ConversationContextMessage | undefined {
  if (remainingCharacters <= 0) return undefined;
  if (message.content.length <= remainingCharacters) return message;
  return { ...message, content: truncateText(message.content, remainingCharacters) };
}

function sumCharacters(messages: ConversationContextMessage[]): number {
  return messages.reduce((total, message) => total + message.content.length, 0);
}

/**
 * Builds a bounded conversational context window.
 *
 * Recent user/assistant turns are retained in chronological order. Older user turns are recalled
 * only when they share exact, normalized tokens with the current query; corrections and explicit
 * preferences win over ordinary older instructions. System and tool messages are never copied.
 * Recalled turns carry both metadata and an in-band lower-trust boundary so a later orchestrator
 * can pass the returned `messages` directly to a provider without treating old text as fresh
 * control instructions.
 */
export function buildConversationContext(
  history: readonly ChatMessage[],
  currentQuery: string,
  options: ConversationContextOptions = {},
): ConversationContext {
  const limits = normalizeOptions(options);
  const eligible: EligibleMessage[] = [];

  history.forEach((message, originalIndex) => {
    if (message.role !== 'user' && message.role !== 'assistant') return;
    const content = sanitizeConversationText(
      typeof message.content === 'string' ? message.content : '',
    );
    if (!content) return;
    eligible.push({ role: message.role, content, originalIndex });
  });

  const recentStart = Math.max(0, eligible.length - limits.recentMessageCount);
  const recentEligible = eligible.slice(recentStart);
  const olderEligible = eligible.slice(0, recentStart);
  const queryText = sanitizeConversationText(currentQuery);
  const queryTokens = conversationTokens(queryText);

  const rankedRecalls = olderEligible
    .map((message) => rankOlderUserMessage(message, queryTokens, limits.minOverlapTokens))
    .filter((message): message is RankedRecall => message !== undefined)
    .sort(compareRecalls)
    .slice(0, limits.maxRecalledMessages);

  const hasCurrentQuery = queryText.length > 0;
  const historyMessageBudget = Math.max(0, limits.maxMessages - (hasCurrentQuery ? 1 : 0));
  const recentToKeep = recentEligible.slice(-Math.min(recentEligible.length, historyMessageBudget));
  const recallBudget = Math.max(0, historyMessageBudget - recentToKeep.length);
  const recallsToKeep = rankedRecalls.slice(0, recallBudget);

  const recentMessages = recentToKeep.map((message) =>
    toRecentMessage(message, limits.maxMessageCharacters),
  );
  const recalledMessages = recallsToKeep.map((message) => ({
    ...message,
    content: wrapRecalledContent(message.content, message.role, limits.maxMessageCharacters),
  }));
  const currentMessage: ConversationContextMessage | undefined = hasCurrentQuery
    ? {
        role: 'user',
        content: truncateText(queryText, limits.maxMessageCharacters),
        source: 'current-query',
        trust: 'current',
      }
    : undefined;

  // Character fitting gives priority to the current query, then the newest recent turns, and
  // finally lower-trust recalls. This means a tight cap cannot evict the request being answered.
  let remainingCharacters = limits.maxCharacters;
  const fittedCurrent = currentMessage
    ? fitMessageToBudget(currentMessage, remainingCharacters)
    : undefined;
  if (fittedCurrent) remainingCharacters -= fittedCurrent.content.length;

  const fittedRecentReversed: ConversationContextMessage[] = [];
  for (let i = recentMessages.length - 1; i >= 0 && remainingCharacters > 0; i -= 1) {
    const fitted = fitMessageToBudget(recentMessages[i]!, remainingCharacters);
    if (!fitted) break;
    fittedRecentReversed.push(fitted);
    remainingCharacters -= fitted.content.length;
  }
  const fittedRecent = fittedRecentReversed.reverse();

  const fittedRecalls: RecalledConversationMessage[] = [];
  for (const recalled of recalledMessages) {
    if (remainingCharacters <= 0) break;
    const fitted = fitMessageToBudget(recalled, remainingCharacters);
    if (!fitted) break;
    fittedRecalls.push(fitted as RecalledConversationMessage);
    remainingCharacters -= fitted.content.length;
  }

  const chronologicalHistory = [...fittedRecent, ...fittedRecalls].sort(
    (left, right) => (left.originalIndex ?? -1) - (right.originalIndex ?? -1),
  );
  const messages = currentMessage
    ? [...chronologicalHistory, fittedCurrent].filter(
        (message): message is ConversationContextMessage => message !== undefined,
      )
    : chronologicalHistory;

  const includedHistoryCount = messages.filter(
    (message) => message.source !== 'current-query',
  ).length;
  return {
    messages,
    recalled: fittedRecalls,
    omittedMessageCount: Math.max(0, eligible.length - includedHistoryCount),
    characterCount: sumCharacters(messages),
  };
}
