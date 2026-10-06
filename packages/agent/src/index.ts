/**
 * @excel-agent/agent
 *
 * The intelligence layer for ExcelAgento. It is UI-free, network-capable, and
 * deterministic-by-default:
 *
 * - `analysis`        deterministic column profiling, search, and intent parsing
 * - `context`         token-efficient sheet context + prompt + model output parser
 * - `tools`           tool catalog derived from the live engine registry
 * - `memory`          self-learning store of verified (query -> operation) pairs
 * - `providers`       BYOK adapters (Groq, OpenRouter, Gemini) with retry/timeout
 * - `orchestrator`    ordered multi-layer control plane with a hard guardrail
 */
export * from './analysis.js';
export * from './context.js';
export * from './evidence.js';
export * from './briefing.js';
export * from './memory.js';
export * from './orchestrator.js';
export * from './providers.js';
export * from './read-tools.js';
export * from './statistical-tools.js';
export * from './tools.js';
export type * from './types.js';
