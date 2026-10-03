import type {
  AgentActivityEvent,
  AgentDecision,
  ChatMessage,
  ExecutionPlan,
  ProposedAction,
  ProviderName,
  StreamCallbacks,
} from '@excel-agent/agent';
import type { Workbook } from '@excel-agent/engine';

import { orchestrator } from './agent-runtime.js';
import { defaultModelFor, isDemoKey } from './settings.js';

export interface LLMConfig {
  provider: ProviderName;
  apiKey: string;
  model?: string;
  baseUrl?: string;
}

export interface AgentResponse {
  message: string;
  thought?: string;
  proposedAction?: ProposedAction;
  plan?: ExecutionPlan;
  source?: AgentDecision['source'];
  guardrail?: AgentDecision['guardrail'];
  trace?: AgentDecision['trace'];
  activities?: AgentActivityEvent[];
  /** Token/latency record, present only when the model layer was invoked. */
  telemetry?: AgentDecision['telemetry'];
}

/**
 * Ask ExcelAgento. Routes through the multi-layer orchestrator with streaming,
 * multi-step plan generation, read tools, and invariant-verified execution.
 */
export async function askExcelAgent(
  userQuery: string,
  workbook: Workbook,
  activeSheetName: string,
  config: LLMConfig | null,
  hasUserUploadedFile: boolean,
  options: {
    conversationHistory?: ChatMessage[];
    callbacks?: StreamCallbacks;
    onActivity?: (activity: AgentActivityEvent) => void;
    signal?: AbortSignal;
  } = {},
): Promise<AgentResponse> {
  const decision = await orchestrator.decide({
    query: userQuery,
    workbook,
    sheetName: activeSheetName,
    signal: options.signal,
    config:
      config && !isDemoKey(config.apiKey)
        ? {
            provider: config.provider,
            apiKey: config.apiKey,
            model: config.model ?? defaultModelFor(config.provider),
            baseUrl: config.baseUrl,
          }
        : null,
    hasUserFile: hasUserUploadedFile,
    conversationHistory: options.conversationHistory,
    callbacks: options.callbacks,
    onActivity: options.onActivity,
  });

  return {
    message: decision.message,
    thought: decision.thought,
    ...(decision.action ? { proposedAction: decision.action } : {}),
    ...(decision.plan ? { plan: decision.plan } : {}),
    source: decision.source,
    ...(decision.guardrail ? { guardrail: decision.guardrail } : {}),
    trace: decision.trace,
    activities: decision.activities,
    ...(decision.telemetry ? { telemetry: decision.telemetry } : {}),
  };
}
