import {
  applyOperation,
  cloneWorkbook,
  type OperationRegistry,
  type Preview,
  type Sheet,
  type Workbook,
} from '@excel-agent/engine';

import { analyzeSpreadsheetIntentAndData, type ProposedAction } from './analysis.js';
import { buildSystemPrompt, parseModelOutput } from './context.js';
import {
  calculateAggregate,
  getWorkbookOverview,
  profileColumn,
  readCellRange,
  READ_TOOL_DEFINITIONS,
  searchSheet,
} from './read-tools.js';
import { complete, completeStream, FALLBACK_MODELS, ProviderError } from './providers.js';
import { buildToolCatalog, type ToolDescriptor } from './tools.js';
import type {
  AgentActivityEvent,
  AgentDecision,
  ChatMessage,
  DecideInput,
  ExecutionPlan,
  ExecutionPlanStep,
  GuardrailReport,
  LlmTelemetry,
  MemoryStore,
  OrchestratorOptions,
  ProviderConfig,
  ProviderResponse,
  ToolDefinition,
  TraceStep,
} from './types.js';

const GREETING = /^(hi|hello|hey|hola|greetings|yo|howdy|sup|namaste|hlo)\b/i;

function isDemoKey(key: string | undefined): boolean {
  if (!key) return true;
  const normalized = key.trim().toLowerCase();
  return normalized.length === 0 || normalized.startsWith('demo');
}

/**
 * ExcelAgento Multi-Layer Intelligent Orchestrator.
 *
 * Tier 0: Fast deterministic & memory cache (offline, zero-token, instant).
 * Tier 1: Interactive Multi-Agent Conductor + Specialists:
 *   - Read-tools for ground-truth data verification without hallucinations.
 *   - Multi-step planning with transactional step-by-step previews.
 *   - Real-time token and thought streaming.
 *   - Automatic provider & model fallback on decommissioned or rate-limited models.
 *   - Guardrail verification (schema -> engine -> invariants -> preview).
 */
export class ExcelAgentOrchestrator {
  private readonly registry: OperationRegistry;
  private readonly memory: MemoryStore | undefined;
  private readonly memorySimilarity: number;
  private readonly catalog: ToolDescriptor[];
  private readonly toolDefinitions: ToolDefinition[];

  constructor(options: OrchestratorOptions) {
    this.registry = options.registry;
    this.memory = options.memory;
    this.memorySimilarity = options.memorySimilarity ?? 0.6;
    this.catalog = buildToolCatalog(options.registry);
    this.toolDefinitions = this.buildAllToolDefinitions();
  }

  get tools(): ToolDescriptor[] {
    return this.catalog;
  }

  private buildAllToolDefinitions(): ToolDefinition[] {
    const definitions: ToolDefinition[] = [...READ_TOOL_DEFINITIONS];

    for (const tool of this.catalog) {
      definitions.push({
        type: 'function',
        function: {
          name: tool.name,
          description: tool.description,
          parameters: (tool.jsonSchema as Record<string, unknown>) ?? {
            type: 'object',
            properties: {},
          },
        },
      });
    }

    definitions.push({
      type: 'function',
      function: {
        name: 'create_execution_plan',
        description:
          'Create an ordered multi-step execution plan when a user request requires multiple operations in sequence.',
        parameters: {
          type: 'object',
          properties: {
            title: { type: 'string', description: 'Brief title for the plan' },
            description: { type: 'string', description: 'Overall summary of the plan' },
            steps: {
              type: 'array',
              items: {
                type: 'object',
                properties: {
                  operation: {
                    type: 'string',
                    description: 'Operation name, e.g. "format_dates", "sort_range"',
                  },
                  args: {
                    type: 'object',
                    description: 'Arguments object matching the operation schema',
                  },
                  description: {
                    type: 'string',
                    description: 'Explanation of what this step achieves',
                  },
                },
                required: ['operation', 'args', 'description'],
              },
              description: 'Ordered sequence of steps to apply',
            },
          },
          required: ['title', 'steps'],
        },
      },
    });

    return definitions;
  }

  async decide(input: DecideInput): Promise<AgentDecision> {
    const trace: TraceStep[] = [];
    const activities: AgentActivityEvent[] = [];

    const emitActivity = (
      type: AgentActivityEvent['type'],
      agent: string,
      summary: string,
      detail?: unknown,
    ) => {
      const event: AgentActivityEvent = {
        id: `act-${Date.now()}-${activities.length}`,
        type,
        agent,
        summary,
        detail,
        timestamp: Date.now(),
      };
      activities.push(event);
      input.onActivity?.(event);
    };

    const sheet =
      input.workbook.sheets.find((candidate) => candidate.name === input.sheetName) ??
      input.workbook.sheets[0];
    const sheetName = sheet?.name ?? input.sheetName;

    // Layer 1 - Intent: Social greetings
    if (GREETING.test(input.query.trim())) {
      const message = !input.hasUserFile
        ? 'Hello! I am **ExcelAgento**, your intelligent enterprise spreadsheet copilot. Upload an Excel or CSV file (or choose a sample fixture), and I can clean data, format dates, normalize text, deduplicate, filter, compute metrics, and execute multi-step operations.'
        : `Hello! I am **ExcelAgento**. **${sheetName}** is loaded with ${
            sheet?.rows.length ?? 0
          } rows. What would you like to inspect, analyze, or transform?`;
      trace.push({ layer: 'intent', summary: 'Greeting detected; no operation proposed.' });
      emitActivity('status', 'Conductor', 'Welcomed user.');
      return { message, source: 'heuristic', trace, activities };
    }

    // Layer 2 - Memory: Replay proven learned operation
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
        emitActivity('status', 'Memory', `Replayed verified operation "${memoryHit.operation}".`);
        return {
          message: `I recognised this as a repeat of a verified action and re-ran it on **${sheetName}**.`,
          action,
          guardrail,
          source: 'memory',
          trace,
          activities,
        };
      }
    }

    return this.plan(input, sheetName, sheet, trace, activities, emitActivity);
  }

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
    activities: AgentActivityEvent[],
    emitActivity: (
      type: AgentActivityEvent['type'],
      agent: string,
      summary: string,
      detail?: unknown,
    ) => void,
  ): Promise<AgentDecision> {
    // Layer 3 - Heuristic Fast Path
    const heuristic = analyzeSpreadsheetIntentAndData(input.query, input.workbook, sheetName);
    trace.push({
      layer: 'heuristic',
      summary: heuristic.proposedAction
        ? `Heuristic planner proposed ${heuristic.proposedAction.name}.`
        : 'Heuristic planner produced an informational answer.',
    });

    let llmAction: ProposedAction | undefined;
    let llmPlan: ExecutionPlan | undefined;
    let llmMessage: string | undefined;
    let llmThought: string | undefined;
    let telemetry: LlmTelemetry | undefined;

    const config = input.config;

    // Layer 4 - LLM Conductor & Specialists
    if (config && !isDemoKey(config.apiKey) && sheet) {
      const startedAt = Date.now();
      emitActivity('thinking', 'Conductor', `Analyzing request for ${sheetName}...`);

      const messages: ChatMessage[] = [
        { role: 'system', content: buildSystemPrompt(sheet, this.catalog) },
      ];

      if (input.conversationHistory && input.conversationHistory.length > 0) {
        const recentHistory = input.conversationHistory.slice(-8);
        for (const historyItem of recentHistory) {
          if (historyItem.role !== 'system') {
            messages.push({
              role: historyItem.role,
              content: historyItem.content,
            });
          }
        }
      }

      messages.push({ role: 'user', content: input.query });

      // Build model retry & fallback list
      const candidateModels = [config.model, ...(FALLBACK_MODELS[config.provider] ?? [])].filter(
        (m, i, arr): m is string => Boolean(m) && arr.indexOf(m) === i,
      );

      let response: ProviderResponse | undefined;
      let lastError: unknown;
      let usedModel = config.model || candidateModels[0] || 'default';

      for (const candidateModel of candidateModels) {
        const currentConfig: ProviderConfig = { ...config, model: candidateModel };
        try {
          if (input.callbacks && typeof input.callbacks.onToken === 'function') {
            response = await completeStream(
              messages,
              currentConfig,
              input.callbacks,
              this.toolDefinitions,
            );
          } else {
            response = await complete(messages, currentConfig, this.toolDefinitions);
          }
          usedModel = candidateModel;
          break;
        } catch (err) {
          lastError = err;
          const msg = err instanceof Error ? err.message : String(err);
          const isQuotaOrModelUnavailable =
            msg.includes('model_decommissioned') ||
            msg.includes('not found') ||
            msg.includes('does not exist') ||
            msg.includes('Request too large') ||
            msg.includes('rate_limit') ||
            (err instanceof ProviderError &&
              (err.status === 404 || err.status === 413 || err.status === 429));

          if (
            isQuotaOrModelUnavailable &&
            candidateModel !== candidateModels[candidateModels.length - 1]
          ) {
            emitActivity(
              'status',
              'Conductor',
              `Model ${candidateModel} quota/rate limit reached; falling back to alternative model...`,
            );
            continue;
          }
          break;
        }
      }

      if (response) {
        llmThought = response.thought;
        let finalResponseContent = response.content;

        // Tool-calling loop: execute read-tools if called by model (max 3 turns)
        let turns = 0;
        let currentToolCalls = response.toolCalls;

        while (currentToolCalls && currentToolCalls.length > 0 && turns < 3) {
          turns += 1;
          const toolCall = currentToolCalls[0]!;
          const fnName = toolCall.function.name;
          let fnArgs: Record<string, unknown> = {};
          try {
            fnArgs = JSON.parse(toolCall.function.arguments || '{}');
          } catch {
            fnArgs = {};
          }

          // Check if this is a read tool
          const isReadTool = [
            'get_workbook_overview',
            'profile_column',
            'read_cell_range',
            'search_sheet',
            'calculate_aggregate',
          ].includes(fnName);

          if (isReadTool) {
            emitActivity(
              'inspecting',
              'Data Analyst',
              `Reading sheet data via ${fnName}...`,
              fnArgs,
            );
            const toolOutput = this.executeReadTool(input.workbook, sheetName, fnName, fnArgs);

            messages.push({
              role: 'assistant',
              content: response.content || '',
              tool_calls: [toolCall],
            });

            messages.push({
              role: 'tool',
              name: fnName,
              tool_call_id: toolCall.id,
              content: JSON.stringify(toolOutput),
            });

            try {
              const followUp = await complete(
                messages,
                { ...config, model: usedModel },
                this.toolDefinitions,
              );
              finalResponseContent = followUp.content;
              currentToolCalls = followUp.toolCalls;
              if (followUp.thought)
                llmThought = (llmThought ? `${llmThought}\n` : '') + followUp.thought;
              continue;
            } catch {
              break;
            }
          }

          // Check if create_execution_plan was called
          if (fnName === 'create_execution_plan') {
            emitActivity(
              'planning',
              'Planner',
              `Formulating execution plan: "${fnArgs.title ?? 'Multi-step update'}"...`,
            );
            const plan = this.buildExecutionPlan(input.workbook, sheetName, fnArgs);
            if (plan) {
              llmPlan = plan;
              llmMessage =
                typeof fnArgs.description === 'string'
                  ? fnArgs.description
                  : `Created a ${plan.steps.length}-step plan to update **${sheetName}**.`;
            }
            break;
          }

          // Check if a single engine write operation was called directly
          if (this.registry.get(fnName)) {
            emitActivity(
              'guardrail_check',
              'Guardrail',
              `Validating proposed operation "${fnName}"...`,
            );
            llmAction = {
              name: fnName,
              args: { ...fnArgs, sheet: fnArgs.sheet ?? sheetName },
              explanation:
                typeof fnArgs.explanation === 'string' ? fnArgs.explanation : `Execute ${fnName}`,
              category: 'transform',
            };
            break;
          }

          break;
        }

        // If no tool call produced an action/plan, fallback to parsing content JSON
        if (!llmAction && !llmPlan && finalResponseContent) {
          const parsed = parseModelOutput(finalResponseContent);
          if (parsed.action) {
            llmAction = {
              name: parsed.action.name,
              args: { ...parsed.action.args, sheet: parsed.action.args.sheet ?? sheetName },
              explanation: parsed.action.explanation,
              category: 'transform',
            };
          }
          if (parsed.message) llmMessage = parsed.message;
        } else if (finalResponseContent && !llmMessage) {
          llmMessage = finalResponseContent;
        }

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
          layer: 'conductor',
          summary: `${response.provider}/${response.model} responded.`,
          detail: {
            proposedOperation: llmAction?.name,
            hasPlan: Boolean(llmPlan),
            totalTokens: telemetry.totalTokens,
          },
          durationMs: telemetry.latencyMs,
        });
      } else {
        const reason = lastError instanceof ProviderError ? lastError.message : String(lastError);
        telemetry = {
          provider: config.provider,
          model: config.model ?? 'unknown',
          latencyMs: Date.now() - startedAt,
          ok: false,
          error: reason,
        };
        trace.push({
          layer: 'conductor',
          summary: `Provider unavailable: ${reason}`,
          durationMs: telemetry.latencyMs,
        });
        emitActivity('status', 'Conductor', `Provider error: ${reason}`);
      }
    }

    // Layer 5 - Guardrail verification for multi-step plan
    if (llmPlan) {
      emitActivity(
        'guardrail_check',
        'Guardrail',
        `Verifying plan: ${llmPlan.steps.length} steps...`,
      );
      return {
        message: llmMessage || `Plan prepared with ${llmPlan.steps.length} steps.`,
        thought: llmThought,
        plan: llmPlan,
        source: 'llm',
        trace,
        activities,
        ...(telemetry ? { telemetry } : {}),
      };
    }

    // Guardrail verification for single action
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
        emitActivity('status', 'Guardrail', `Approved operation "${candidate.name}".`);
        return {
          message: llmMessage || heuristic.message,
          thought: llmThought,
          action: candidate,
          guardrail,
          source: llmAction && candidate === llmAction ? 'llm' : 'heuristic',
          trace,
          activities,
          ...(telemetry ? { telemetry } : {}),
        };
      }
    }

    // Conversational fallback
    trace.push({ layer: 'guardrail', summary: 'Informational answer; no mutation proposed.' });
    return {
      message: llmMessage || heuristic.message,
      thought: llmThought,
      source: 'fallback',
      trace,
      activities,
      ...(telemetry ? { telemetry } : {}),
    };
  }

  private executeReadTool(
    workbook: Workbook,
    sheetName: string,
    name: string,
    args: Record<string, unknown>,
  ): unknown {
    switch (name) {
      case 'get_workbook_overview':
        return getWorkbookOverview(workbook);
      case 'profile_column':
        return profileColumn(
          workbook,
          typeof args.sheet === 'string' ? args.sheet : sheetName,
          String(args.column ?? 'A'),
        );
      case 'read_cell_range':
        return readCellRange(
          workbook,
          typeof args.sheet === 'string' ? args.sheet : sheetName,
          typeof args.startRow === 'number' ? args.startRow : 1,
          typeof args.endRow === 'number' ? args.endRow : 15,
          typeof args.startColumn === 'string' ? args.startColumn : 'A',
          typeof args.endColumn === 'string' ? args.endColumn : undefined,
        );
      case 'search_sheet':
        return searchSheet(
          workbook,
          typeof args.sheet === 'string' ? args.sheet : sheetName,
          String(args.query ?? ''),
          typeof args.limit === 'number' ? args.limit : 15,
        );
      case 'calculate_aggregate':
        return calculateAggregate(
          workbook,
          typeof args.sheet === 'string' ? args.sheet : sheetName,
          String(args.column ?? 'A'),
          args.metric as 'sum' | 'avg' | 'min' | 'max' | 'count' | 'count_distinct',
        );
      default:
        return { error: `Unknown read tool "${name}".` };
    }
  }

  private buildExecutionPlan(
    workbook: Workbook,
    sheetName: string,
    args: Record<string, unknown>,
  ): ExecutionPlan | undefined {
    const rawSteps = Array.isArray(args.steps) ? args.steps : [];
    if (rawSteps.length === 0) return undefined;

    let simWorkbook = cloneWorkbook(workbook);
    const steps: ExecutionPlanStep[] = [];
    let totalAffected = 0;

    for (let i = 0; i < rawSteps.length; i += 1) {
      const rawStep = rawSteps[i] as {
        operation?: string;
        args?: Record<string, unknown>;
        description?: string;
      };
      if (!rawStep || typeof rawStep.operation !== 'string') continue;

      const opName = rawStep.operation;
      const stepArgs = { ...(rawStep.args ?? {}), sheet: rawStep.args?.sheet ?? sheetName };

      const stepAction: ProposedAction = {
        name: opName,
        args: stepArgs,
        explanation: rawStep.description || `Step ${i + 1}: ${opName}`,
        category: 'transform',
      };

      const guardrail = this.guardrail(simWorkbook, stepAction);
      const stepPreview = guardrail.preview;

      if (guardrail.passed) {
        const execRes = applyOperation(simWorkbook, opName, stepArgs, { registry: this.registry });
        if (execRes.ok) {
          simWorkbook = execRes.workbook;
          totalAffected += execRes.report.affectedCells;
        }
      }

      steps.push({
        id: `step-${Date.now()}-${i}`,
        operation: opName,
        args: stepArgs,
        description: rawStep.description || `Execute ${opName}`,
        category: 'transform',
        status: guardrail.passed ? 'pending' : 'error',
        preview: stepPreview,
        error: guardrail.errors.length > 0 ? guardrail.errors.join('; ') : undefined,
      });
    }

    if (steps.length === 0) return undefined;

    return {
      id: `plan-${Date.now()}`,
      title: typeof args.title === 'string' ? args.title : 'Spreadsheet Execution Plan',
      description:
        typeof args.description === 'string' ? args.description : 'Multi-step transformation plan',
      steps,
      status: 'pending',
      totalAffectedCells: totalAffected,
    };
  }

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
