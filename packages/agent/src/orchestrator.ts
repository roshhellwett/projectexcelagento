import {
  cloneWorkbook,
  type OperationRegistry,
  type Preview,
  type Sheet,
  type Workbook,
} from '@excel-agent/engine';

import { analyzeSpreadsheetIntentAndData, type ProposedAction } from './analysis.js';
import { buildSystemPrompt, parseModelOutput } from './context.js';
import { complete, ProviderError } from './providers.js';
import { buildToolCatalog, type ToolDescriptor } from './tools.js';
import type {
  AgentDecision,
  DecideInput,
  GuardrailReport,
  LlmTelemetry,
  MemoryStore,
  OrchestratorOptions,
  TraceStep,
} from './types.js';

const GREETING = /^(hi|hello|hey|hola|greetings|yo|howdy|sup|namaste|hlo)\b/i;

function isDemoKey(key: string | undefined): boolean {
  if (!key) return true;
  const normalized = key.trim().toLowerCase();
  return normalized.length === 0 || normalized.startsWith('demo');
}

/**
 * ExcelAgento's control plane. Each conversational turn passes through ordered
 * layers so a model can never mutate the workbook without schema, engine, and
 * preview agreement:
 *
 *   intent -> memory -> heuristic -> llm -> guardrail -> (verification on execute)
 */
export class ExcelAgentOrchestrator {
  private readonly registry: OperationRegistry;

  private readonly memory: MemoryStore | undefined;

  private readonly memorySimilarity: number;

  private readonly catalog: ToolDescriptor[];

  constructor(options: OrchestratorOptions) {
    this.registry = options.registry;
    this.memory = options.memory;
    this.memorySimilarity = options.memorySimilarity ?? 0.6;
    this.catalog = buildToolCatalog(options.registry);
  }

  get tools(): ToolDescriptor[] {
    return this.catalog;
  }

  async decide(input: DecideInput): Promise<AgentDecision> {
    const trace: TraceStep[] = [];
    const sheet =
      input.workbook.sheets.find((candidate) => candidate.name === input.sheetName) ??
      input.workbook.sheets[0];
    const sheetName = sheet?.name ?? input.sheetName;

    // Layer 1 - Intent: social turns never touch the workbook.
    if (GREETING.test(input.query.trim())) {
      const message = !input.hasUserFile
        ? 'Hello. I am **ExcelAgento**. Upload an Excel or CSV file (or pick a sample fixture) and I can normalize dates, clean text, deduplicate, sort, filter, and compute aggregates.'
        : `Hello. I am **ExcelAgento**. **${sheetName}** is loaded with ${
            sheet?.rows.length ?? 0
          } rows. Tell me what to change and I will show you a verified preview first.`;
      trace.push({ layer: 'intent', summary: 'Greeting detected; no operation proposed.' });
      return { message, source: 'heuristic', trace };
    }

    // Layer 2 - Memory: replay a previously verified operation.
    const memoryHit = this.memory?.retrieve(input.query, sheetName, this.memorySimilarity);
    if (memoryHit) {
      const action: ProposedAction = {
        name: memoryHit.operation,
        args: { ...memoryHit.args },
        explanation: 'Reusing a previously verified action for a similar request.',
        category: 'transform',
      };
      const guardrail = this.guardrail(input.workbook, action);
      trace.push({
        layer: 'memory',
        summary: `Memory hit on "${memoryHit.rawQuery}" (${memoryHit.successes} ok / ${memoryHit.failures} failed).`,
        detail: { operation: memoryHit.operation, passed: guardrail.passed },
      });
      if (guardrail.passed) {
        return {
          message: `I recognised this as a repeat of a verified action and re-ran it on **${sheetName}**.`,
          action,
          guardrail,
          source: 'memory',
          trace,
        };
      }
    }

    return this.plan(input, sheetName, sheet, trace);
  }

  /**
   * Self-learning hook. The host calls this after an action has been applied (or
   * failed) so proven associations are reinforced and replayable next time.
   */
  learn(outcome: {
    query: string;
    sheetName: string;
    operation: string;
    args: Record<string, unknown>;
    success: boolean;
  }): void {
    if (!this.memory) return;
    if (outcome.success) {
      this.memory.remember({
        key: outcome.query,
        rawQuery: outcome.query,
        operation: outcome.operation,
        args: outcome.args,
        sheetName: outcome.sheetName,
      });
    }
    this.memory.recordOutcome(outcome.operation, outcome.sheetName, outcome.success);
  }

  private async plan(
    input: DecideInput,
    sheetName: string,
    sheet: Sheet | undefined,
    trace: TraceStep[],
  ): Promise<AgentDecision> {
    // Layer 3 - Heuristic: deterministic, offline-first understanding.
    const heuristic = analyzeSpreadsheetIntentAndData(input.query, input.workbook, sheetName);
    trace.push({
      layer: 'heuristic',
      summary: heuristic.proposedAction
        ? `Heuristic planner proposed ${heuristic.proposedAction.name}.`
        : 'Heuristic planner produced an informational answer.',
    });

    // Layer 4 - LLM: optional BYOK reasoning layer.
    let llmAction: ProposedAction | undefined;
    let llmMessage: string | undefined;
    let telemetry: LlmTelemetry | undefined;
    const config = input.config;
    if (config && !isDemoKey(config.apiKey) && sheet) {
      const startedAt = Date.now();
      try {
        const response = await complete(
          [
            { role: 'system', content: buildSystemPrompt(sheet, this.catalog) },
            { role: 'user', content: input.query },
          ],
          config,
        );
        const parsed = parseModelOutput(response.content);
        if (parsed.action) {
          llmAction = {
            name: parsed.action.name,
            args: { ...parsed.action.args, sheet: parsed.action.args.sheet ?? sheetName },
            explanation: parsed.action.explanation,
            category: 'transform',
          };
        }
        if (parsed.message) llmMessage = parsed.message;
        const summed =
          (response.usage?.promptTokens ?? 0) + (response.usage?.completionTokens ?? 0);
        telemetry = {
          provider: response.provider,
          model: response.model,
          promptTokens: response.usage?.promptTokens,
          completionTokens: response.usage?.completionTokens,
          totalTokens: response.usage?.totalTokens ?? (summed > 0 ? summed : undefined),
          latencyMs: Date.now() - startedAt,
          ok: true,
        };
        trace.push({
          layer: 'llm',
          summary: `${response.provider}/${response.model} responded.`,
          detail: { proposedOperation: llmAction?.name, totalTokens: telemetry.totalTokens },
          durationMs: telemetry.latencyMs,
        });
      } catch (error) {
        const reason = error instanceof ProviderError ? error.message : String(error);
        telemetry = {
          provider: config.provider,
          model: config.model ?? 'unknown',
          latencyMs: Date.now() - startedAt,
          ok: false,
          error: reason,
        };
        trace.push({
          layer: 'llm',
          summary: `Provider unavailable: ${reason}`,
          durationMs: telemetry.latencyMs,
        });
      }
    }

    // Layer 5 - Guardrail: every candidate must satisfy schema + engine + preview.
    const candidates: ProposedAction[] = [];
    if (llmAction) candidates.push(llmAction);
    if (heuristic.proposedAction) candidates.push(heuristic.proposedAction);

    for (const candidate of candidates) {
      const guardrail = this.guardrail(input.workbook, candidate);
      trace.push({
        layer: 'guardrail',
        summary: `${candidate.name}: ${guardrail.passed ? 'passed' : 'blocked'}.`,
        detail: guardrail.errors,
      });
      if (guardrail.passed) {
        return {
          message: llmMessage || heuristic.message,
          action: candidate,
          guardrail,
          source: llmAction && candidate === llmAction ? 'llm' : 'heuristic',
          trace,
          ...(telemetry ? { telemetry } : {}),
        };
      }
    }

    // Fallback - answer conversationally, never silently guessing an operation.
    trace.push({ layer: 'guardrail', summary: 'No candidate action passed the guardrail.' });
    return {
      message: llmMessage || heuristic.message,
      source: 'fallback',
      trace,
      ...(telemetry ? { telemetry } : {}),
    };
  }

  /** Validate an action against the registry, the engine, and a bounded preview. */
  guardrail(workbook: Workbook, action: ProposedAction): GuardrailReport {
    const operation = this.registry.get(action.name);
    if (!operation) {
      return {
        passed: false,
        schemaValid: false,
        engineValid: false,
        errors: [`Unknown operation "${action.name}".`],
        warnings: [],
        requiresConfirmation: false,
      };
    }

    const parsed = operation.schema.safeParse(action.args);
    if (!parsed.success) {
      return {
        passed: false,
        schemaValid: false,
        engineValid: false,
        errors: parsed.error.issues.map((issue) => `${issue.path.join('.')}: ${issue.message}`),
        warnings: [],
        requiresConfirmation: false,
      };
    }

    const before = cloneWorkbook(workbook);
    try {
      const validation = operation.validate(cloneWorkbook(before), parsed.data);
      if (!validation.valid) {
        return {
          passed: false,
          schemaValid: true,
          engineValid: false,
          errors: validation.errors.map((issue) => issue.message),
          warnings: [],
          requiresConfirmation: false,
        };
      }
      const preview: Preview = operation.preview(cloneWorkbook(before), parsed.data);
      const warnings = [...validation.warnings, ...preview.warnings].map((issue) => issue.message);
      return {
        passed: preview.valid,
        schemaValid: true,
        engineValid: true,
        preview,
        errors: preview.errors.map((issue) => issue.message),
        warnings,
        requiresConfirmation: preview.requiresConfirmation || preview.affectedCells > 200,
      };
    } catch (error) {
      return {
        passed: false,
        schemaValid: true,
        engineValid: false,
        errors: [error instanceof Error ? error.message : 'Preview failed.'],
        warnings: [],
        requiresConfirmation: false,
      };
    }
  }
}

export function createOrchestrator(options: OrchestratorOptions): ExcelAgentOrchestrator {
  return new ExcelAgentOrchestrator(options);
}
