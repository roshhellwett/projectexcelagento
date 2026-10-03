import type { Sheet } from '@excel-agent/engine';

import { getColumnProfiles } from './analysis.js';
import { describeTools, type ToolDescriptor } from './tools.js';

/**
 * A compact, token-efficient profile of the active sheet. Full cell data is never
 * sent to the model — only headers, types, cardinality, aggregates, and a sample.
 */
export function buildSheetContext(sheet: Sheet, sampleSize = 3): string {
  const rowCount = sheet.rows.length;
  const colCount = Math.max(...sheet.rows.map((row) => row.length), 0);
  const profiles = getColumnProfiles(sheet);

  const columns = profiles.map((profile) => ({
    col: profile.letter,
    name: profile.rawName,
    type: profile.isNumeric ? 'numeric' : profile.isDate ? 'date' : 'text',
    nonBlank: profile.nonBlankCount,
    distinct: profile.distinct.size,
    samples: Array.from(profile.distinct.keys()).slice(0, 3),
    ...(profile.isNumeric && profile.sum !== undefined
      ? {
          sum: Math.round(profile.sum * 100) / 100,
          avg: Math.round((profile.avg ?? 0) * 100) / 100,
          min: profile.min,
          max: profile.max,
        }
      : {}),
  }));

  const sampleRows = sheet.rows.slice(0, sampleSize).map((row, rowIndex) => {
    const record: Record<string, unknown> = { _row: rowIndex + 1 };
    profiles.forEach((profile, columnIndex) => {
      record[`${profile.letter}_${profile.rawName}`] = row[columnIndex]?.value ?? null;
    });
    return record;
  });

  return JSON.stringify({
    sheet: sheet.name,
    rows: rowCount,
    cols: colCount,
    columns,
    sampleRows,
  });
}

export function buildSystemPrompt(sheet: Sheet, catalog: ToolDescriptor[]): string {
  return `You are ExcelAgento, a senior data analyst and spreadsheet copilot.

Active worksheet:
${buildSheetContext(sheet)}

Rules:
1. Understand any phrasing, including English, Hindi, Hinglish, slang, or shorthand.
2. Ground every number in the worksheet profile above. Never invent data.
3. When the user wants a change, emit exactly one JSON object inside a \`\`\`json fence with shape {"name": "<operation>", "args": {...}, "explanation": "..."}.
4. Column letters must match the worksheet. Use the sheet name "${sheet.name}".
5. If you are only answering a question, reply in prose with no JSON block.

Available operations:
${describeTools(catalog)}`;
}

export interface ParsedModelOutput {
  message: string;
  action?: { name: string; args: Record<string, unknown>; explanation: string };
}

/** Balanced-brace scan for the first top-level JSON object containing "name", string-aware. */
function extractBalancedJson(content: string): string | null {
  for (let start = 0; start < content.length; start += 1) {
    if (content[start] !== '{') continue;
    let depth = 0;
    let inString = false;
    let escaped = false;
    for (let i = start; i < content.length; i += 1) {
      const ch = content[i];
      if (inString) {
        if (escaped) escaped = false;
        else if (ch === '\\') escaped = true;
        else if (ch === '"') inString = false;
        continue;
      }
      if (ch === '"') inString = true;
      else if (ch === '{') depth += 1;
      else if (ch === '}') {
        depth -= 1;
        if (depth === 0) {
          const candidate = content.slice(start, i + 1);
          if (candidate.includes('"name"')) return candidate;
          break;
        }
      }
    }
  }
  return null;
}

/** Extract the first JSON action block from a model reply, tolerating fences and prose. */
export function parseModelOutput(content: string): ParsedModelOutput {
  const candidates: string[] = [];
  const fenced = content.match(/```(?:json)?\s*([\s\S]*?)```/g);
  if (fenced) {
    for (const block of fenced) {
      candidates.push(
        block
          .replace(/```(?:json)?/g, '')
          .replace(/```/g, '')
          .trim(),
      );
    }
  }
  const bare = extractBalancedJson(content);
  if (bare) candidates.push(bare);

  for (const candidate of candidates) {
    try {
      const parsed = JSON.parse(candidate) as {
        name?: unknown;
        args?: unknown;
        explanation?: unknown;
      };
      if (
        typeof parsed.name === 'string' &&
        typeof parsed.args === 'object' &&
        parsed.args !== null &&
        !Array.isArray(parsed.args)
      ) {
        return {
          message: stripJsonBlocks(content),
          action: {
            name: parsed.name,
            args: parsed.args as Record<string, unknown>,
            explanation:
              typeof parsed.explanation === 'string'
                ? parsed.explanation
                : `Execute ${parsed.name}.`,
          },
        };
      }
    } catch {
      // Not valid JSON; try the next candidate.
    }
  }

  return { message: content.trim() };
}

function stripJsonBlocks(content: string): string {
  const cleaned = content
    .replace(/```(?:json)?\s*[\s\S]*?```/g, '')
    .replace(/\{[\s\S]*?"name"[\s\S]*?\}/g, '')
    .trim();
  return cleaned || content.trim();
}
