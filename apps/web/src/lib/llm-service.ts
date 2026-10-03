import type { AgentDecision, ProposedAction, ProviderName } from '@excel-agent/agent';
import type { Workbook } from '@excel-agent/engine';

import { orchestrator } from './agent-runtime.js';
import { defaultModelFor, isDemoKey } from './settings.js';

export interface LLMConfig {
  provider: ProviderName;
  apiKey: string;
  model?: string;
}

export interface AgentResponse {
  message: string;
  proposedAction?: ProposedAction;
  source?: AgentDecision['source'];
  guardrail?: AgentDecision['guardrail'];
  trace?: AgentDecision['trace'];
}

/**
 * Ask ExcelAgento. Every call routes through the multi-layer orchestrator: the
 * deterministic planner is the ground truth, the BYOK model is an optional
 * reasoning layer, and each candidate action must clear the guardrail first.
 */
export async function askExcelAgent(
  userQuery: string,
  workbook: Workbook,
  activeSheetName: string,
  config: LLMConfig | null,
  hasUserUploadedFile: boolean,
): Promise<AgentResponse> {
  const decision = await orchestrator.decide({
    query: userQuery,
    workbook,
    sheetName: activeSheetName,
    config:
      config && !isDemoKey(config.apiKey)
        ? {
            provider: config.provider,
            apiKey: config.apiKey,
            model: config.model ?? defaultModelFor(config.provider),
          }
        : null,
    hasUserFile: hasUserUploadedFile,
  });

  return {
    message: decision.message,
    ...(decision.action ? { proposedAction: decision.action } : {}),
    source: decision.source,
    ...(decision.guardrail ? { guardrail: decision.guardrail } : {}),
    trace: decision.trace,
  };
}
