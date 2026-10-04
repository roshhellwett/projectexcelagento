import type { OperationRegistry, Preview, Workbook } from '@excel-agent/engine';

import type { ProposedAction } from './analysis.js';

/** Supported BYOK inference providers. */
export type ProviderName = 'groq' | 'openrouter' | 'gemini' | 'openai' | 'custom';

export interface ProviderConfig {
  provider: ProviderName;
  apiKey: string;
  model?: string;
  temperature?: number;
  maxTokens?: number;
  /** Custom endpoint base URL for Ollama, LM Studio, or self-hosted OpenAI proxies. */
  baseUrl?: string;
  /** Hard timeout for a single inference request. Defaults to 30s. */
  timeoutMs?: number;
  /** Number of retries on transient (429/5xx/network) failures. Defaults to 2. */
  retries?: number;
  signal?: AbortSignal;
}

export interface ToolCall {
  id: string;
  type: 'function';
  function: {
    name: string;
    arguments: string;
  };
}

export interface ToolDefinition {
  type: 'function';
  function: {
    name: string;
    description: string;
    parameters: Record<string, unknown>;
  };
}

export interface ChatMessage {
  role: 'system' | 'user' | 'assistant' | 'tool';
  content: string;
  name?: string;
  tool_call_id?: string;
  tool_calls?: ToolCall[];
  thought?: string;
}

export interface StreamCallbacks {
  onToken?: (token: string) => void;
  onThinking?: (thought: string) => void;
  onToolCall?: (toolCall: ToolCall) => void;
}

export interface ProviderResponse {
  content: string;
  thought?: string;
  toolCalls?: ToolCall[];
  provider: ProviderName;
  model: string;
  usage?: { promptTokens?: number; completionTokens?: number; totalTokens?: number };
}

export interface ProviderAdapter {
  readonly name: ProviderName;
  readonly defaultModel: string;
  complete(
    messages: ChatMessage[],
    config: ProviderConfig,
    tools?: ToolDefinition[],
  ): Promise<ProviderResponse>;
  completeStream?(
    messages: ChatMessage[],
    config: ProviderConfig,
    callbacks: StreamCallbacks,
    tools?: ToolDefinition[],
  ): Promise<ProviderResponse>;
}

/** Machine-readable cost/latency record for one BYOK inference attempt. */
export interface LlmTelemetry {
  provider: ProviderName;
  /** The model the provider reports having actually served the request. */
  model: string;
  promptTokens?: number;
  completionTokens?: number;
  totalTokens?: number;
  latencyMs: number;
  ok: boolean;
  error?: string;
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
  | 'intent'
  | 'memory'
  | 'heuristic'
  | 'conductor'
  | 'planner'
  | 'specialist'
  | 'llm'
  | 'guardrail'
  | 'verification';

export interface TraceStep {
  layer: AgentLayerName;
  summary: string;
  detail?: unknown;
  durationMs?: number;
}

export type DecisionSource = 'memory' | 'heuristic' | 'llm' | 'fallback';

export interface ExecutionPlanStep {
  id: string;
  operation: string;
  args: Record<string, unknown>;
  description: string;
  category?: ProposedAction['category'];
  status?: 'pending' | 'executing' | 'completed' | 'error';
  preview?: Preview;
  error?: string;
}

export interface ExecutionPlan {
  id: string;
  title: string;
  description: string;
  steps: ExecutionPlanStep[];
  status: 'pending' | 'applied' | 'error';
  totalAffectedCells?: number;
}

export type AgentEventType =
  | 'thinking'
  | 'inspecting'
  | 'planning'
  | 'tool_call'
  | 'guardrail_check'
  /** Something was refused or had to be given up. Must read as distinct from progress. */
  | 'warning'
  | 'status';

export interface AgentActivityEvent {
  id: string;
  type: AgentEventType;
  agent: string;
  summary: string;
  detail?: unknown;
  timestamp: number;
}

/** The orchestrator's answer for one conversational turn. */
export interface AgentDecision {
  message: string;
  thought?: string;
  action?: ProposedAction;
  plan?: ExecutionPlan;
  guardrail?: GuardrailReport;
  insights?: string[];
  source: DecisionSource;
  trace: TraceStep[];
  activities?: AgentActivityEvent[];
  /** Present only when the BYOK model layer was actually invoked. */
  telemetry?: LlmTelemetry;
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
  /**
   * Identity of the sheet shape this was learned against: its header names in order.
   *
   * Learned arguments are column LETTERS, and letters move. Learn "sort by revenue" on a sheet
   * whose revenue is column E, then delete column D, and the replay still targets E - which now
   * holds a different field. Schema validation passes, the guardrail approves, and the user's
   * data is silently reshuffled. Comparing this fingerprint before a replay is what stops it.
   */
  schemaFingerprint?: string;
}

export interface MemoryStore {
  remember(
    record: Omit<MemoryRecord, 'id' | 'createdAt' | 'lastUsedAt' | 'successes' | 'failures'>,
  ): MemoryRecord;
  recordOutcome(operation: string, sheetName: string, success: boolean, key?: string): void;
  retrieve(
    query: string,
    sheetName: string,
    threshold?: number,
  ): Promise<MemoryRecord | undefined> | MemoryRecord | undefined;
  /**
   * Retrieves a record only if its learned arguments still fit the sheet's current shape.
   * Callers replaying a learned action against real user data must use this, not `retrieve`.
   */
  retrieveForWorkbook?(
    query: string,
    workbook: Workbook,
    sheetName: string,
    threshold?: number,
  ): Promise<MemoryRecord | undefined> | MemoryRecord | undefined;
  entries(): MemoryRecord[];
  toJSON(): string;
  confidenceOf?(record: MemoryRecord): number;
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
  /** Multi-turn conversation history for interactive context. */
  conversationHistory?: ChatMessage[];
  /** Streaming token / thought callbacks for real-time interactive response. */
  callbacks?: StreamCallbacks;
  /** Real-time activity callback for live subagent and tool feedback. */
  onActivity?: (activity: AgentActivityEvent) => void;
  /** Aborts the in-flight provider request (used by the stop button). */
  signal?: AbortSignal;
};
