import type { OperationRegistry, Preview, Workbook } from '@excel-agent/engine';

import type { ProposedAction } from './analysis.js';

/** Supported BYOK inference providers. */
export type ProviderName = 'groq' | 'openrouter' | 'gemini';

export interface ProviderConfig {
  provider: ProviderName;
  apiKey: string;
  model?: string;
  temperature?: number;
  maxTokens?: number;
  /** Hard timeout for a single inference request. Defaults to 30s. */
  timeoutMs?: number;
  /** Number of retries on transient (429/5xx/network) failures. Defaults to 2. */
  retries?: number;
  signal?: AbortSignal;
}

export interface ChatMessage {
  role: 'system' | 'user' | 'assistant';
  content: string;
}

export interface ProviderResponse {
  content: string;
  provider: ProviderName;
  model: string;
  usage?: { promptTokens?: number; completionTokens?: number };
}

export interface ProviderAdapter {
  readonly name: ProviderName;
  readonly defaultModel: string;
  complete(messages: ChatMessage[], config: ProviderConfig): Promise<ProviderResponse>;
}

/** Machine-checkable result of the guardrail layer. */
export interface GuardrailReport {
  passed: boolean;
  schemaValid: boolean;
  engineValid: boolean;
  preview?: Preview;
  errors: string[];
  warnings: string[];
  requiresConfirmation: boolean;
}

export type AgentLayerName =
  'intent' | 'memory' | 'heuristic' | 'llm' | 'guardrail' | 'verification';

export interface TraceStep {
  layer: AgentLayerName;
  summary: string;
  detail?: unknown;
  durationMs?: number;
}

export type DecisionSource = 'memory' | 'heuristic' | 'llm' | 'fallback';

/** The orchestrator's answer for one conversational turn. */
export interface AgentDecision {
  message: string;
  action?: ProposedAction;
  guardrail?: GuardrailReport;
  insights?: string[];
  source: DecisionSource;
  trace: TraceStep[];
}

/** A single learned (query -> operation) association. Enables the self-learning layer. */
export interface MemoryRecord {
  id: string;
  /** Normalized query used for retrieval. */
  key: string;
  rawQuery: string;
  operation: string;
  args: Record<string, unknown>;
  sheetName: string;
  successes: number;
  failures: number;
  createdAt: number;
  lastUsedAt: number;
}

export interface MemoryStore {
  remember(
    record: Omit<MemoryRecord, 'id' | 'createdAt' | 'lastUsedAt' | 'successes' | 'failures'>,
  ): MemoryRecord;
  recordOutcome(operation: string, sheetName: string, success: boolean): void;
  retrieve(query: string, sheetName: string, threshold?: number): MemoryRecord | undefined;
  entries(): MemoryRecord[];
  toJSON(): string;
  clear(): void;
}

export interface OrchestratorOptions {
  registry: OperationRegistry;
  memory?: MemoryStore;
  /** Confidence at/above which a memory hit is used without asking the LLM. */
  memoryConfidence?: number;
  /** Minimum similarity for fuzzy memory retrieval. */
  memorySimilarity?: number;
}

export type DecideInput = {
  query: string;
  workbook: Workbook;
  sheetName: string;
  config?: ProviderConfig | null;
  /** True once the user has loaded their own file (affects greeting copy). */
  hasUserFile?: boolean;
};
