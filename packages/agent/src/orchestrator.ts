import {
  applyOperation,
  cloneWorkbook,
  columnToIndex,
  indexToColumn,
  type OperationRegistry,
  type Preview,
  type Sheet,
  type Workbook,
} from '@excel-agent/engine';

import { analyzeSpreadsheetIntentAndData, type ProposedAction } from './analysis.js';
import { buildSystemPrompt, parseModelOutput } from './context.js';
import {
  complete,
  completeStream,
  defaultModelFor,
  ProviderError,
  throwIfCancelled,
} from './providers.js';
import { InferenceUsage } from './inference-usage.js';
import { evidenceFromToolResult } from './evidence.js';
import {
  executeWorkbookReadTool,
  isWorkbookReadTool,
  untrustedToolOutput,
  WORKBOOK_READ_TOOLS,
} from './read-tool-runtime.js';
import { sheetFingerprint } from './memory.js';
import { buildToolCatalog, type ToolDescriptor } from './tools.js';
import { isComplexRequest, runMultiAgentTurn } from './multi-agent.js';
import { analyzeReconciliationIntent, isReconciliationRequest } from './reconciliation-intent.js';
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
  private readonly memoryConfidence: number;
  private readonly catalog: ToolDescriptor[];
  private readonly toolDefinitions: ToolDefinition[];

  constructor(options: OrchestratorOptions) {
    this.registry = options.registry;
    this.memory = options.memory;
    this.memorySimilarity = options.memorySimilarity ?? 0.6;
    this.memoryConfidence = options.memoryConfidence ?? 0;
    this.catalog = buildToolCatalog(options.registry);
    this.toolDefinitions = this.buildAllToolDefinitions();
  }

  get tools(): ToolDescriptor[] {
    return this.catalog;
  }

  private saveWorkingStep(
    input: DecideInput,
    stepType: string,
    payload: Record<string, unknown>,
  ): void {
    if (!input.sessionId || !this.memory?.saveWorkingMemory) {
      return;
    }
    void this.memory.saveWorkingMemory(input.sessionId, stepType, payload);
  }

  private buildAllToolDefinitions(): ToolDefinition[] {
    const definitions: ToolDefinition[] = [...WORKBOOK_READ_TOOLS];

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
              maxItems: 25,
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
    const signal = input.signal ?? input.config?.signal;
    const provider = input.config?.provider ?? 'groq';
    try {
      throwIfCancelled(provider, signal);
      const decision = await this.decideTurn(input);
      throwIfCancelled(provider, signal);
      return decision;
    } catch (error) {
      if (signal?.aborted) {
        return {
          message: 'Cancelled. Nothing was changed.',
          source: 'fallback',
          trace: [{ layer: 'intent', summary: 'Cancelled by the caller; no action was proposed.' }],
          activities: [],
        };
      }
      throw error;
    }
  }

  private async decideTurn(input: DecideInput): Promise<AgentDecision> {
    const trace: TraceStep[] = [];
    const activities: AgentActivityEvent[] = [];

    const emitActivity = (
      type: AgentActivityEvent['type'],
      agent: string,
      summary: string,
      detail?: unknown,
      tokens?: AgentActivityEvent['tokens'],
    ) => {
      const extractedTokens =
        tokens ??
        (detail && typeof detail === 'object' && 'tokens' in detail
          ? (detail as { tokens?: AgentActivityEvent['tokens'] }).tokens
          : undefined);
      const event: AgentActivityEvent = {
        id: `act-${Date.now()}-${activities.length}`,
        type,
        agent,
        summary,
        detail,
        tokens: extractedTokens,
        timestamp: Date.now(),
      };
      activities.push(event);
      input.onActivity?.(event);
    };

    const sheet =
      input.workbook.sheets.find((candidate) => candidate.name === input.sheetName) ??
      input.workbook.sheets[0];
    const sheetName = sheet?.name ?? input.sheetName;

    this.saveWorkingStep(input, 'turn_started', {
      query: input.query,
      sheetName,
      rowCount: sheet?.rows.length ?? 0,
      colCount: sheet?.rows[0]?.length ?? 0,
      timestamp: Date.now(),
    });

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

    // Cross-sheet business jobs must resolve fresh bindings and tolerance; one-sheet memory
    // fingerprints cannot authorize replay against a different right-side schema.
    if (analyzeReconciliationIntent(input.query, input.workbook))
      return this.plan(input, sheetName, sheet, trace, activities, emitActivity);

    // Layer 2 - Memory: Replay proven learned operation
    const memoryHit = this.memory?.retrieveForWorkbook
      ? await this.memory.retrieveForWorkbook(
          input.query,
          input.workbook,
          sheetName,
          this.memorySimilarity,
        )
      : await this.memory?.retrieve(input.query, sheetName, this.memorySimilarity);
    if (memoryHit) {
      const confidentEnough = this.memory?.confidenceOf
        ? this.memory.confidenceOf(memoryHit) >= this.memoryConfidence
        : true;
      if (!confidentEnough) {
        trace.push({
          layer: 'memory',
          summary: 'Memory hit skipped: below memoryConfidence threshold.',
        });
        return this.plan(input, sheetName, sheet, trace, activities, emitActivity);
      }
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
    /**
     * The sheet shape the outcome applies to. Recorded with the memory entry so a later replay
     * can be refused if the columns have since moved. Omitting it stores an entry that can never
     * be replayed, which is safe but wastes the learning.
     */
    workbook?: Workbook;
  }): void {
    if (!this.memory) return;
    if (outcome.success) {
      const fingerprint = outcome.workbook
        ? sheetFingerprint(outcome.workbook, outcome.sheetName)
        : undefined;
      this.memory.remember({
        key: outcome.query,
        rawQuery: outcome.query,
        operation: outcome.operation,
        args: outcome.args,
        sheetName: outcome.sheetName,
        ...(fingerprint === undefined ? {} : { schemaFingerprint: fingerprint }),
      });
    }
    this.memory.recordOutcome(outcome.operation, outcome.sheetName, outcome.success, outcome.query);
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
      tokens?: AgentActivityEvent['tokens'],
    ) => void,
  ): Promise<AgentDecision> {
    const CONVERSATIONAL_FOLLOWUP_PATTERN =
      /^(?:perform(?:\s+then|\s+it)?|do\s+it(?:\s+then)?|run\s+it(?:\s+then)?|proceed(?:\s+then)?|go\s+ahead|yes|ok|sure|please\s+do|please\s+perform|execute(?:\s+it|\s+then)?|apply(?:\s+it|\s+then)?|clean\s+it(?:\s+then)?|structure\s+it(?:\s+then)?|do\s+that(?:\s+then)?|start(?:\s+then)?|count\s+it|show\s+me|give\s+me\s+(?:the\s+)?data\s*\??|show\s+(?:me\s+)?(?:the\s+)?data\s*\??|tell\s+me\s*\??|what\s+is\s+it\s*\??|what\s+does\s+it\s+say\s*\??|what\s+happend\s*\??|what\s+happened\s*\??|what\s+task\s+i\s+gave\s+you\s*\??|\?)$/i;

    const isConversationalFollowup = CONVERSATIONAL_FOLLOWUP_PATTERN.test(input.query.trim());

    let effectiveQuery = input.query;
    if (
      isConversationalFollowup &&
      input.conversationHistory &&
      input.conversationHistory.length > 0
    ) {
      for (let i = input.conversationHistory.length - 1; i >= 0; i--) {
        const h = input.conversationHistory[i];
        if (
          h &&
          h.role === 'user' &&
          h.content.trim().length > 1 &&
          !CONVERSATIONAL_FOLLOWUP_PATTERN.test(h.content.trim())
        ) {
          effectiveQuery = h.content;
          break;
        }
      }
    }

    // Layer 3 - Heuristic Fast Path
    const heuristic = analyzeSpreadsheetIntentAndData(effectiveQuery, input.workbook, sheetName);
    const evidence = [...(heuristic.evidence ?? [])];
    trace.push({
      layer: 'heuristic',
      summary: heuristic.proposedAction
        ? `Heuristic planner proposed ${heuristic.proposedAction.name}.`
        : heuristic.clarification
          ? `Heuristic requested clarification: ${heuristic.clarification.question}`
          : 'Heuristic planner produced an informational answer.',
    });

    if (!isComplexRequest(effectiveQuery) && heuristic.clarification) {
      emitActivity(
        'status',
        'Conductor',
        'Proactively asking clarification questions to avoid errors.',
      );
      return {
        message: heuristic.message,
        clarification: heuristic.clarification,
        source: 'heuristic',
        trace,
        activities,
      };
    }

    if (
      !isComplexRequest(effectiveQuery) &&
      isReconciliationRequest(effectiveQuery) &&
      analyzeReconciliationIntent(effectiveQuery, input.workbook)
    ) {
      if (!heuristic.proposedAction)
        return {
          message: heuristic.message,
          clarification: heuristic.clarification,
          evidence,
          source: 'heuristic',
          trace,
          activities,
        };
      const guardrail = this.guardrail(input.workbook, heuristic.proposedAction);
      trace.push({
        layer: 'guardrail',
        summary: guardrail.passed
          ? 'Reconciliation job verified against both source sheets.'
          : 'Reconciliation job refused safely.',
      });
      emitActivity(
        guardrail.passed ? 'status' : 'warning',
        'Verifier',
        guardrail.passed
          ? 'Report bindings verified. Source sheets remain unchanged; review before generating the workbook.'
          : guardrail.errors.join(' '),
      );
      return {
        message: guardrail.passed
          ? heuristic.message
          : `I could not safely prepare this reconciliation report. ${guardrail.errors.join(' ')} Nothing was changed.`,
        action: guardrail.passed ? heuristic.proposedAction : undefined,
        guardrail,
        evidence,
        source: 'heuristic',
        trace,
        activities,
      };
    }

    // Layer 3.5 - Multi-agent: complex requests are decomposed, executed, and reviewed.
    // Every layer after this one is the single-agent fast path.
    if (input.config && sheet && isComplexRequest(input.query) && !isDemoKey(input.config.apiKey)) {
      try {
        const multi = await runMultiAgentTurn(
          {
            query: input.query,
            workbook: input.workbook,
            sheetName,
            config: input.signal ? { ...input.config, signal: input.signal } : input.config,
            emit: emitActivity,
            callbacks: input.callbacks,
          },
          { registry: this.registry },
        );
        trace.push(...multi.trace);
        trace.push({
          layer: 'conductor',
          summary: 'Complex request handled by the multi-agent pipeline.',
        });
        return {
          ...multi,
          trace,
          activities: [...activities, ...(multi.activities ?? [])],
          evidence: multi.evidence,
          telemetry: multi.telemetry,
        };
      } catch (error) {
        throwIfCancelled(input.config.provider, input.signal ?? input.config.signal);
        // A failed specialist turn must never throw from decide(); offer the
        // single-agent / heuristic path as a continued fallback the user can see.
        trace.push({
          layer: 'conductor',
          summary: 'Multi-agent pipeline failed; continuing to the single-agent path.',
          detail: error instanceof Error ? error.message : String(error),
        });
      }
    }

    let llmAction: ProposedAction | undefined;
    let llmPlan: ExecutionPlan | undefined;
    let llmMessage: string | undefined;
    let llmThought: string | undefined;
    let telemetry: LlmTelemetry | undefined;
    let modelResponded = false;

    const config = input.config
      ? input.signal
        ? { ...input.config, signal: input.signal }
        : input.config
      : input.config;


    // Layer 4 - LLM Conductor & Specialists
    if (config && !isDemoKey(config.apiKey) && sheet) {
      const startedAt = Date.now();
      emitActivity('thinking', 'Conductor', `Analyzing request for ${sheetName}...`);
      input.callbacks?.onThinking?.(
        `Conductor: Analyzing request for "${sheetName}" (${sheet?.rows.length ?? 0} rows). Grounding context with workbook schema...\n`,
      );

      const messages: ChatMessage[] = [
        { role: 'system', content: buildSystemPrompt(sheet, this.catalog, input.workbook) },
      ];

      if (this.memory?.getWorkingMemory && input.sessionId) {
        try {
          const workingSteps = await this.memory.getWorkingMemory(input.sessionId);
          if (Array.isArray(workingSteps) && workingSteps.length > 0) {
            const recentSteps = workingSteps.slice(-8);
            const memorySummary = recentSteps
              .map((s) => {
                const item = s as { step_type?: string; payload?: Record<string, unknown> };
                const stepType = item.step_type || 'step';
                const p = item.payload || {};
                if (stepType === 'read_tool_inspection') {
                  return `- Prior check (${p.tool}): ${p.summary || JSON.stringify(p.args)}`;
                }
                if (stepType === 'decision_approved') {
                  return `- Prior executed operation: ${p.operation}`;
                }
                if (stepType === 'turn_completed') {
                  return `- Prior answer given: "${p.message ? String(p.message).slice(0, 160) : ''}"`;
                }
                return `- Prior step [${stepType}]: ${JSON.stringify(p).slice(0, 140)}`;
              })
              .join('\n');

            if (memorySummary) {
              messages.push({
                role: 'system',
                content: `Working session memory from Supabase (context across previous turns in this session):\n${memorySummary}`,
              });
            }
          }
        } catch {
          // Non-blocking fallback
        }
      }

      if (input.conversationHistory && input.conversationHistory.length > 0) {
        const recentHistory = input.conversationHistory.slice(-30);
        for (const historyItem of recentHistory) {
          if (historyItem.role !== 'system') {
            messages.push({
              role: historyItem.role,
              content: historyItem.content,
            });
          }
        }
      }

      if (isConversationalFollowup && effectiveQuery !== input.query) {
        messages.push({
          role: 'system',
          content: `The user said "${input.query}". They are confirming your previous response and want you to fulfill their underlying request: "${effectiveQuery}". Execute the search/query tool immediately and provide the exact answer directly.`,
        });
      }

      // Detect migration / population intent between sheets
      const isMigrationOrPopulation =
        /\b(?:add\s+(?:all\s+)?data\s+here|migrate\s+(?:the\s+)?data|populate\s+(?:the\s+)?(?:data|sheet|table)|transfer\s+(?:the\s+)?data|put\s+(?:the\s+)?data\s+here|transfer\s+it|move\s+the\s+data|fill\s+this\s+sheet)\b/i.test(
          input.query.trim(),
        ) ||
        (/\btransfer\b/i.test(input.query) && /\b(?:to|into)\b/i.test(input.query));

      if (isMigrationOrPopulation && input.workbook.sheets.length > 1) {
        let sourceSheetCandidate = input.workbook.sheets.find(
          (s) => s.name !== sheetName && s.rows.length > 3,
        );
        let destSheetCandidate = sheet;

        for (const s of input.workbook.sheets) {
          const normName = s.name.toLowerCase().replace(/[_\s]+/g, ' ');
          const normQ = input.query.toLowerCase().replace(/[_\s]+/g, ' ');
          if (s.name !== sheetName && normQ.includes(normName)) {
            destSheetCandidate = s;
            sourceSheetCandidate = sheet;
            break;
          }
        }

        if (sourceSheetCandidate && destSheetCandidate) {
          const destHeaders = (destSheetCandidate.rows[0] ?? [])
            .map((c) => String(c?.value ?? '').trim())
            .filter(Boolean);
          messages.push({
            role: 'system',
            content: `Data Migration Directive:
The user is requesting: "${input.query}".
- Source Sheet: "${sourceSheetCandidate.name}" (${sourceSheetCandidate.rows.length} rows)
- Destination Sheet: "${destSheetCandidate.name}" (${destSheetCandidate.rows.length} rows, target headers: [${destHeaders.join(', ')}])
Execute this workflow:
1. Inspect the source sheet records via read_cell_range (e.g. read the top data rows or unparsed text).
2. Extract the structured records that align with the destination headers.
3. Call append_rows with { sheet: "${destSheetCandidate.name}", rows: [...] } to populate the destination sheet. Do NOT refuse or state that the source sheet contains code or unformatted text; extract the records and insert them into the destination.`,
          });
        }
      }

      messages.push({
        role: 'user',
        content:
          isConversationalFollowup && effectiveQuery !== input.query
            ? `${input.query} (Proceed with task: "${effectiveQuery}")`
            : input.query,
      });

      let response: ProviderResponse | undefined;
      let lastError: unknown;
      const usedModel = config.model?.trim() || defaultModelFor(config.provider);

      const activeMaxTokens =
        config.provider === 'groq'
          ? Math.min(config.maxTokens ?? 8192, 8192)
          : (config.maxTokens ?? 32768);

      const currentConfig: ProviderConfig = {
        ...config,
        model: usedModel,
        maxTokens: activeMaxTokens,
      };

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
      } catch (err) {
        lastError = err;
        const msg = err instanceof Error ? err.message : String(err);

        // A cancellation is the user saying stop, not a provider problem. It must leave the
        // turn immediately: falling through to the heuristic candidate would let a Stop
        // press still hand back an "Apply Changes" card for a request the user abandoned.
        if (input.config?.signal?.aborted || /cancel(?:led|ed) by caller/i.test(msg)) {
          emitActivity('warning', 'Conductor', 'Request cancelled; no action was proposed.');
          trace.push({ layer: 'llm', summary: 'Cancelled by the caller before a decision.' });
          return {
            message: 'Cancelled. Nothing was changed.',
            source: 'fallback',
            trace,
            activities,
          };
        }
      }

      if (response) {
        const usage = new InferenceUsage();
        usage.add(response);
        throwIfCancelled(config.provider, config.signal);
        llmThought = response.thought;
        let finalResponseContent = response.content;

        // Tool-calling loop: execute all tool calls per turn until an action/plan is produced or the turn budget is spent
        let turns = 0;
        let currentToolCalls = response.toolCalls;
        let resolved = false;

        while (!resolved && currentToolCalls && currentToolCalls.length > 0 && turns < 30) {
          throwIfCancelled(config.provider, config.signal);
          turns += 1;
          messages.push({
            role: 'assistant',
            content: finalResponseContent || response.content || '',
            tool_calls: currentToolCalls,
          });

          let executedReadTools = false;

          for (const toolCall of currentToolCalls) {
            const fnName = toolCall.function.name;
            let fnArgs: Record<string, unknown> = {};
            try {
              fnArgs = JSON.parse(toolCall.function.arguments || '{}');
            } catch {
              fnArgs = {};
            }

            // Check if this is a read tool
            const isReadTool = isWorkbookReadTool(fnName);

            if (isReadTool) {
              executedReadTools = true;
              const approxPromptTokens = Math.max(
                20,
                Math.round(messages.reduce((acc, m) => acc + (m.content?.length || 0), 0) / 3.8),
              );
              emitActivity(
                'inspecting',
                'Data Analyst',
                `Reading sheet data via ${fnName}...`,
                fnArgs,
                { promptTokens: approxPromptTokens, totalTokens: approxPromptTokens },
              );
              const argsSummary =
                Object.keys(fnArgs).length > 0 ? ` (${JSON.stringify(fnArgs)})` : '';
              input.callbacks?.onThinking?.(
                `\n[Data Scientist] Reading sheet data via "${fnName}"${argsSummary}...\n`,
              );
              input.callbacks?.onTokenCount?.({
                promptTokens: approxPromptTokens,
                totalTokens: approxPromptTokens,
              });
              const toolOutput = await this.executeReadTool(
                input.workbook,
                sheetName,
                fnName,
                fnArgs,
              );
              const toolEvidence = evidenceFromToolResult(fnName, fnArgs, toolOutput);
              if (toolEvidence && evidence.length < 12) evidence.push(toolEvidence);
              let observationText = '';
              if (typeof toolOutput === 'object' && toolOutput !== null) {
                if (Array.isArray(toolOutput)) {
                  observationText = `Retrieved ${toolOutput.length} record(s).`;
                } else {
                  const outputRecord = toolOutput as Record<string, unknown>;
                  if (
                    'aggregates' in outputRecord &&
                    typeof outputRecord.aggregates === 'object' &&
                    outputRecord.aggregates !== null
                  ) {
                    observationText = `Computed aggregates: ${JSON.stringify(outputRecord.aggregates)}`;
                  } else if ('values' in outputRecord && Array.isArray(outputRecord.values)) {
                    observationText = `Extracted ${outputRecord.values.length} cell value(s).`;
                  } else {
                    observationText = `Inspected schema with keys: ${Object.keys(outputRecord).slice(0, 5).join(', ')}.`;
                  }
                }
              } else {
                observationText = String(toolOutput).slice(0, 120);
              }
              input.callbacks?.onThinking?.(`[Data Scientist] Observation: ${observationText}\n`);
              this.saveWorkingStep(input, 'read_tool_inspection', {
                tool: fnName,
                args: fnArgs,
                summary: `Inspected ${sheetName} via ${fnName}`,
                timestamp: Date.now(),
              });
              messages.push({
                role: 'tool',
                name: fnName,
                tool_call_id: toolCall.id,
                // Read results quote whole rows, so they carry the same injection risk as the
                // sheet profile. Every string in the payload is neutralized before the model
                // ever sees it, and the wrapper states plainly that the contents are data.
                content: untrustedToolOutput(toolOutput),
              });
              const updatedTokens = Math.max(
                approxPromptTokens,
                Math.round(messages.reduce((acc, m) => acc + (m.content?.length || 0), 0) / 3.8),
              );
              input.callbacks?.onTokenCount?.({
                promptTokens: updatedTokens,
                totalTokens: updatedTokens,
              });
              continue;
            }

            // Check if create_execution_plan was called
            if (fnName === 'create_execution_plan') {
              emitActivity(
                'planning',
                'Planner',
                `Formulating execution plan: "${fnArgs.title ?? 'Multi-step update'}"...`,
              );
              input.callbacks?.onThinking?.(
                `\n[Planner] Formulating execution plan: "${fnArgs.title ?? 'Multi-step update'}"...\n`,
              );
              const plan = this.buildExecutionPlan(input.workbook, sheetName, fnArgs);
              if (plan) {
                llmPlan = plan;
                llmMessage =
                  typeof fnArgs.description === 'string'
                    ? fnArgs.description
                    : `Created a ${plan.steps.length}-step plan to update **${sheetName}**.`;
                this.saveWorkingStep(input, 'plan_created', {
                  title: fnArgs.title,
                  steps: plan.steps.map((s) => s.operation),
                  timestamp: Date.now(),
                });
                input.callbacks?.onThinking?.(
                  `[Sentinel] Verifying ${plan.steps.length} operation(s) against mathematical invariants and schema safety...\n`,
                );
              }
              resolved = true;
              break;
            }

            // Check if a single engine write operation was called directly
            if (this.registry.get(fnName)) {
              emitActivity(
                'guardrail_check',
                'Guardrail',
                `Validating proposed operation "${fnName}"...`,
              );
              input.callbacks?.onThinking?.(
                `\n[Sentinel] Validating proposed operation "${fnName}" against engine invariants...\n`,
              );
              llmAction = {
                name: fnName,
                args: { ...fnArgs, sheet: fnArgs.sheet ?? sheetName },
                explanation:
                  typeof fnArgs.explanation === 'string' ? fnArgs.explanation : `Execute ${fnName}`,
                category: 'transform',
              };
              resolved = true;
              break;
            }

            // Unknown tool: stop the loop
            resolved = true;
            break;
          }

          if (resolved) break;
          if (!executedReadTools) break;

          let followUp: ProviderResponse | undefined;
          try {
            if (input.callbacks && typeof input.callbacks.onToken === 'function') {
              followUp = await completeStream(
                messages,
                { ...config, model: usedModel, maxTokens: activeMaxTokens },
                input.callbacks,
                this.toolDefinitions,
              );
            } else {
              followUp = await complete(
                messages,
                { ...config, model: usedModel, maxTokens: activeMaxTokens },
                this.toolDefinitions,
              );
            }
            usage.add(followUp);
            throwIfCancelled(config.provider, config.signal);
          } catch (err) {
            throwIfCancelled(config.provider, config.signal);
            const errMsg = err instanceof Error ? err.message : String(err);
            emitActivity('warning', 'Data Analyst', `Model error: ${errMsg}`);
            return {
              message: `⚠️ **OpenRouter / Model Error (${usedModel}):** ${errMsg}\n\nPlease check your OpenRouter credits, rate limits, or model ID in **Settings** (top right) so you can fix what happened.`,
              source: 'llm',
              trace,
              activities,
              telemetry: {
                provider: config.provider,
                model: config.model ?? 'unknown',
                latencyMs: Date.now() - startedAt,
                ok: false,
                error: errMsg,
              },
            };
          }

          if (!followUp) {
            break;
          }

          finalResponseContent = followUp.content;
          currentToolCalls = followUp.toolCalls;
          if (followUp.thought) {
            llmThought = (llmThought ? `${llmThought}\n` : '') + followUp.thought;
            input.callbacks?.onThinking?.(`\n${followUp.thought}\n`);
          }

          // If the model produced text with no tool calls, check if it's an incomplete thought or self-directed preamble
          const isIncomplete =
            /(?:let\s+me|i\s+will|i'll\s+now|let's|hold\s+on|truncated\s+at|need\s+to\s+pull|verifying|checking|let\s+me\s+actually|cross-check)/i.test(
              finalResponseContent,
            ) ||
            /(?:shows?|indicates?|follows?|following|results?|summary|breakdown|details?|here\s+is|here's\s+what)\s*[:.]?$/i.test(
              finalResponseContent.trim(),
            ) ||
            finalResponseContent.trim().endsWith('—') ||
            finalResponseContent.trim().endsWith('...') ||
            finalResponseContent.trim().endsWith(':') ||
            (executedReadTools &&
              finalResponseContent.trim().length < 280 &&
              /(?:i\s+read|i\s+inspected|i\s+checked|here's\s+what|what\s+the\s+data\s+shows)/i.test(
                finalResponseContent,
              ) &&
              !/\d+(?:\.\d+)?%|\b(?:probability|profit|loss|margin|average|sum|total|ratio)\b.*?\d+/i.test(
                finalResponseContent,
              ));

          if (isIncomplete && (!currentToolCalls || currentToolCalls.length === 0) && turns < 29) {
            messages.push({
              role: 'assistant',
              content: finalResponseContent,
            });
            messages.push({
              role: 'user',
              content:
                'Please proceed immediately without pausing: deliver your complete analytical findings, exact numbers, probabilities, calculations, and final answer directly to the user right now.',
            });

            try {
              let continuation: ProviderResponse;
              if (input.callbacks && typeof input.callbacks.onToken === 'function') {
                continuation = await completeStream(
                  messages,
                  { ...config, model: usedModel, maxTokens: activeMaxTokens },
                  input.callbacks,
                  this.toolDefinitions,
                );
              } else {
                continuation = await complete(
                  messages,
                  { ...config, model: usedModel, maxTokens: activeMaxTokens },
                  this.toolDefinitions,
                );
              }
              usage.add(continuation);
              throwIfCancelled(config.provider, config.signal);
              finalResponseContent = continuation.content;
              currentToolCalls = continuation.toolCalls;
              if (continuation.thought) {
                llmThought = (llmThought ? `${llmThought}\n` : '') + continuation.thought;
                input.callbacks?.onThinking?.(`\n${continuation.thought}\n`);
              }
            } catch (err) {
              throwIfCancelled(config.provider, config.signal);
              const errMsg = err instanceof Error ? err.message : String(err);
              emitActivity('warning', 'Data Analyst', `Continuation error: ${errMsg}`);
            }
          }
        }

        // An unfinished model answer is a failure to answer, never permission to invent results.
        if (
          !llmAction &&
          !llmPlan &&
          finalResponseContent &&
          /(?:here's\s+what\s+the\s+data\s+shows|what\s+the\s+data\s+shows|here's\s+what\s+the\s+numbers\s+show)\s*[:.]?$/i.test(
            finalResponseContent.trim(),
          )
        ) {
          finalResponseContent =
            'The model returned an incomplete analytical answer. I could not verify a result for this request. Please ask for a specific calculation or column; nothing was changed.';
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
        modelResponded = true;

        telemetry = {
          provider: response.provider,
          model: response.model,
          ...usage.snapshot(),
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
          summary: `Model failed: ${reason}`,
          durationMs: telemetry.latencyMs,
        });
        emitActivity('warning', 'Conductor', `Model error: ${reason}`);

        return {
          message: `⚠️ **OpenRouter / Model Error (${usedModel}):** ${reason}\n\nPlease check your OpenRouter credits, rate limits, or model ID in **Settings** (top right) so you can fix what happened.`,
          source: 'llm',
          trace,
          activities,
          telemetry,
        };
      }
    }

    // Layer 5 - Guardrail verification for multi-step plan
    if (llmPlan) {
      emitActivity(
        'guardrail_check',
        'Guardrail',
        `Verifying plan: ${llmPlan.steps.length} steps...`,
      );
      const failedSteps = llmPlan.steps.filter((s) => s.status === 'error').length;
      this.saveWorkingStep(input, 'plan_prepared', {
        title: llmPlan.title,
        steps: llmPlan.steps.map((s) => s.operation),
        status: llmPlan.status,
        timestamp: Date.now(),
      });
      return {
        message:
          llmPlan.status === 'error'
            ? `I couldn't verify any step of that plan: ${llmPlan.steps
                .map((s) => s.error)
                .filter(Boolean)
                .join('; ')}`
            : llmMessage ||
              (failedSteps > 0
                ? `Plan prepared with ${llmPlan.steps.length} steps (${failedSteps} flagged with validation errors).`
                : `Plan prepared with ${llmPlan.steps.length} steps.`),
        thought: llmThought,
        plan: llmPlan,
        evidence,
        source: 'llm',
        trace,
        activities,
        ...(telemetry ? { telemetry } : {}),
      };
    }

    if (!modelResponded && !llmPlan && isComplexRequest(effectiveQuery)) {
      const compositePlan = this.synthesizeCompositePlan(input.workbook, sheetName, effectiveQuery);
      if (compositePlan) {
        emitActivity(
          'planning',
          'Planner',
          `Formulated ${compositePlan.steps.length}-step deliverable plan.`,
        );
        trace.push({
          layer: 'planner',
          summary: `Synthesized verified ${compositePlan.steps.length}-step deliverable plan: ${compositePlan.title}.`,
        });
        return {
          message: `I analyzed your workbook and prepared the deliverable: **${compositePlan.title}**.\n\n${compositePlan.description}\n\nReview the ${compositePlan.steps.length} verified steps below and click **Apply Plan** to execute.`,
          plan: compositePlan,
          source: 'heuristic',
          trace,
          activities,
        };
      }
    }

    // Guardrail verification for single action
    const candidates: ProposedAction[] = [];
    const isCritiqueOrFeedback =
      /\b(?:copy\s*pasted?|duplicate\s*sheet|why\s+did\s+you|wrong|error|mistake|nothing|broken|undo|revert|why\s+llms?|not\s+thinking|not\s+understanding)\b/i.test(
        input.query,
      );
    const isSheetOrFilterQuery =
      /(?:filter|extract|separate|new sheet|separate sheet|create sheet|duplicate sheet|delete sheet)/i.test(
        input.query,
      );
    const isInformationalLlmProposal = llmAction?.name === 'aggregate_column';

    if (modelResponded) {
      // When an LLM model is engaged, the LLM is the brain.
      if (isSheetOrFilterQuery && isInformationalLlmProposal && heuristic.proposedAction) {
        candidates.push(heuristic.proposedAction);
        if (llmAction) candidates.push(llmAction);
      } else if (llmAction) {
        candidates.push(llmAction);
      } else if (!isCritiqueOrFeedback && heuristic.proposedAction) {
        // Fallback for explicit operational directives where model emitted text without a tool call,
        // but NEVER hijack turns where the user is expressing critique, feedback, or conversational reasoning.
        candidates.push(heuristic.proposedAction);
      }
    } else {
      // Offline mode without model or LLM disabled: use heuristic candidate
      if (heuristic.proposedAction) {
        candidates.push(heuristic.proposedAction);
      }
    }

    let blockedByGuardrail: { action: ProposedAction; guardrail: GuardrailReport } | null = null;

    for (const candidate of candidates) {
      const guardrail = this.guardrail(input.workbook, candidate);
      trace.push({
        layer: 'guardrail',
        summary: `${candidate.name}: ${guardrail.passed ? 'passed' : 'blocked'}.`,
        detail: guardrail.errors,
      });
      if (guardrail.passed) {
        // Prevent useless 0-affected-cell mutations from being proposed to the user
        if (
          guardrail.preview &&
          guardrail.preview.affectedCells === 0 &&
          candidate.category !== 'filter'
        ) {
          emitActivity(
            'status',
            'Guardrail',
            `Skipped "${candidate.name}" because 0 cells would be affected.`,
          );
          continue;
        }

        emitActivity('status', 'Guardrail', `Approved operation "${candidate.name}".`);
        const isLlmCandidate = llmAction !== undefined && candidate === llmAction;
        this.saveWorkingStep(input, 'decision_approved', {
          action: candidate.name,
          source: isLlmCandidate ? 'llm' : 'heuristic',
          timestamp: Date.now(),
        });
        const defaultActionMsg = `I've prepared the **${candidate.name}** operation for **${sheetName}**: ${candidate.explanation}. Click **Apply Changes** to proceed.`;
        return {
          message: isLlmCandidate ? llmMessage?.trim() || defaultActionMsg : heuristic.message,
          thought: isLlmCandidate ? llmThought : undefined,
          action: candidate,
          evidence,
          guardrail,
          source: isLlmCandidate ? 'llm' : 'heuristic',
          trace,
          activities,
          ...(telemetry ? { telemetry } : {}),
        };
      }
      blockedByGuardrail ??= { action: candidate, guardrail };
    }

    // An action was proposed and refused. Say so plainly rather than silently answering a
    // question the user did not ask, which is indistinguishable from "the agent decided not to
    // act" and leaves the user with no idea anything was even attempted.
    if (blockedByGuardrail) {
      const reasons = blockedByGuardrail.guardrail.errors.length
        ? blockedByGuardrail.guardrail.errors
        : ['the guardrail could not verify it'];
      emitActivity(
        'warning',
        'Guardrail',
        `Blocked "${blockedByGuardrail.action.name}" - ${reasons[0]}`,
      );
      trace.push({
        layer: 'guardrail',
        summary: `Refused to substitute a different action for the blocked "${blockedByGuardrail.action.name}".`,
        detail: reasons,
      });
      return {
        message: `I could not safely apply **${blockedByGuardrail.action.name}**. ${reasons.join(' ')} Nothing has been changed - adjust the request or run the operation manually.`,
        thought: llmThought,
        source: 'fallback',
        trace,
        activities,
        ...(telemetry ? { telemetry } : {}),
      };
    }

    // Conversational fallback
    trace.push({ layer: 'guardrail', summary: 'Informational answer; no mutation proposed.' });
    const finalMsg = llmMessage?.trim() || heuristic.message;
    this.saveWorkingStep(input, 'turn_completed', {
      source: llmMessage ? 'llm' : 'heuristic',
      message: finalMsg,
      timestamp: Date.now(),
    });
    return {
      message: finalMsg,
      thought: llmThought,
      evidence,
      source: llmMessage ? 'llm' : 'fallback',
      trace,
      activities,
      ...(telemetry ? { telemetry } : {}),
    };
  }

  private async executeReadTool(
    workbook: Workbook,
    sheetName: string,
    name: string,
    args: Record<string, unknown>,
  ): Promise<unknown> {
    return executeWorkbookReadTool(workbook, sheetName, name, args);
  }

  private buildExecutionPlan(
    workbook: Workbook,
    sheetName: string,
    args: Record<string, unknown>,
  ): ExecutionPlan | undefined {
    const rawSteps = Array.isArray(args.steps) ? args.steps : [];
    if (rawSteps.length === 0 || rawSteps.length > 25) return undefined;

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
      const stepArgs = {
        ...(rawStep.args ?? {}),
        ...(rawStep.operation === 'reconcile_sheets'
          ? {}
          : { sheet: rawStep.args?.sheet ?? sheetName }),
      };

      const stepAction: ProposedAction = {
        name: opName,
        args: stepArgs,
        explanation: rawStep.description || `Step ${i + 1}: ${opName}`,
        category: 'transform',
      };

      const guardrail = this.guardrail(simWorkbook, stepAction);
      const stepPreview = guardrail.preview;
      const errors = [...guardrail.errors];

      if (guardrail.passed) {
        const execRes = applyOperation(simWorkbook, opName, stepArgs, {
          registry: this.registry,
          confirmed: true,
        });
        if (execRes.ok) {
          simWorkbook = execRes.workbook;
          totalAffected += execRes.report.affectedCells;
        } else {
          errors.push(...execRes.error.messages);
        }
      }

      steps.push({
        id: `step-${Date.now()}-${i}`,
        operation: opName,
        args: stepArgs,
        description: rawStep.description || `Execute ${opName}`,
        category: 'transform',
        status: guardrail.passed && errors.length === 0 ? 'pending' : 'error',
        preview: stepPreview,
        error: errors.length > 0 ? errors.join('; ') : undefined,
      });
    }

    if (steps.length === 0) return undefined;

    const failed = steps.filter((s) => s.status === 'error').length;
    const status = failed > 0 ? 'error' : 'pending';

    return {
      id: `plan-${Date.now()}`,
      title: typeof args.title === 'string' ? args.title : 'Spreadsheet Execution Plan',
      description:
        typeof args.description === 'string' ? args.description : 'Multi-step transformation plan',
      steps,
      status,
      totalAffectedCells: totalAffected,
    };
  }

  private synthesizeCompositePlan(
    workbook: Workbook,
    activeSheetName: string,
    query: string,
  ): ExecutionPlan | undefined {
    const q = query.toLowerCase();
    const isCompositeOutcome =
      /\b(?:sales\s+report|prepare|reconcil|discrepanc|summariz|by\s+region|clean)\b/i.test(q) &&
      isComplexRequest(query);

    if (!isCompositeOutcome) return undefined;

    const sheets = workbook.sheets;
    if (sheets.length === 0) return undefined;

    const ordersSheet =
      sheets.find((s) => /order|sale|transact/i.test(s.name)) ??
      sheets.find((s) => s.name === activeSheetName) ??
      sheets[0]!;

    const invoicesSheet =
      sheets.find((s) => /invoice|bill/i.test(s.name) && s.name !== ordersSheet.name) ??
      sheets.find((s) => s.name !== ordersSheet.name);

    const steps: Array<{ operation: string; args: Record<string, unknown>; description: string }> =
      [];

    const findCol = (sheet: Sheet, pattern: RegExp): string | undefined => {
      const headerRow = sheet.rows[0] ?? [];
      for (let i = 0; i < headerRow.length; i++) {
        const val = String(headerRow[i]?.value ?? '');
        if (pattern.test(val)) return indexToColumn(i);
      }
      return undefined;
    };

    // Stage 1: Clean orders if requested
    if (/\b(?:clean|deduplicat|normaliz|format)\b/i.test(q)) {
      const orderIdCol = findCol(ordersSheet, /order\s*id|id|order\s*#|ref/i) ?? 'A';
      const seen = new Set<string>();
      let hasDuplicates = false;
      const colIdx = columnToIndex(orderIdCol) ?? 0;
      for (let r = 1; r < ordersSheet.rows.length; r++) {
        const val = String(ordersSheet.rows[r]?.[colIdx]?.value ?? '');
        if (val && seen.has(val)) {
          hasDuplicates = true;
          break;
        }
        if (val) seen.add(val);
      }

      if (hasDuplicates) {
        steps.push({
          operation: 'delete_duplicates',
          args: { sheet: ordersSheet.name, columns: [orderIdCol], keep: 'first', headerRow: 1 },
          description: `Clean ${ordersSheet.name}: deduplicate records keyed on ${orderIdCol}`,
        });
      } else {
        const textCols: string[] = [];
        (ordersSheet.rows[0] ?? []).forEach((cell, idx) => {
          const val = String(cell?.value ?? '');
          if (/name|customer|region|status|type/i.test(val)) {
            textCols.push(indexToColumn(idx));
          }
        });
        if (textCols.length > 0) {
          steps.push({
            operation: 'normalize_text',
            args: {
              sheet: ordersSheet.name,
              columns: textCols,
              trim: true,
              collapseWhitespace: true,
              headerRow: 1,
            },
            description: `Clean ${ordersSheet.name}: normalize and trim whitespace in text columns (${textCols.join(', ')})`,
          });
        }
      }
    }

    // Stage 2: Reconcile orders against invoices
    if (/\b(?:reconcil|discrepanc|match|against\s+invoices?)\b/i.test(q) && invoicesSheet) {
      const leftKey = findCol(ordersSheet, /order\s*id|id|order\s*#|invoice\s*id/i) ?? 'A';
      const rightKey = findCol(invoicesSheet, /order\s*id|invoice\s*id|id|order\s*#/i) ?? 'A';
      const leftAmount = findCol(ordersSheet, /amount|total|sales|price/i);
      const rightAmount = findCol(invoicesSheet, /amount|total|invoiced|due/i);

      steps.push({
        operation: 'reconcile_sheets',
        args: {
          leftSheet: ordersSheet.name,
          rightSheet: invoicesSheet.name,
          leftKeys: [leftKey],
          rightKeys: [rightKey],
          ...(leftAmount && rightAmount ? { leftAmount, rightAmount } : {}),
          tolerance: 0,
          reportPrefix: 'Sales Reconciliation',
        },
        description: `Reconcile ${ordersSheet.name} against ${invoicesSheet.name} and generate exception reports`,
      });
    }

    // Stage 3: Summarize performance by region
    if (/\b(?:summariz|perform|region|group|rollup)\b/i.test(q)) {
      const regionCol = findCol(ordersSheet, /region|territory|zone|country|state/i);
      const amountCol = findCol(ordersSheet, /amount|sales|total|revenue/i);

      if (regionCol && amountCol) {
        steps.push({
          operation: 'group_and_summarize',
          args: {
            sheet: ordersSheet.name,
            groupBy: [regionCol],
            valueColumn: amountCol,
            aggregation: 'sum',
            targetSheet: 'Regional Performance',
          },
          description: `Summarize sales performance grouped by region (${regionCol}) into Regional Performance sheet`,
        });
      }
    }

    if (steps.length < 2) return undefined;

    return this.buildExecutionPlan(workbook, activeSheetName, {
      title: 'Executive Monthly Sales & Reconciliation Deliverable',
      description: `End-to-end autonomous analysis: cleaned ${ordersSheet.name}, reconciled against ${invoicesSheet?.name ?? 'invoices'}, flagged discrepancies, and summarized regional KPIs into presentation-ready sheets.`,
      steps,
    });
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
