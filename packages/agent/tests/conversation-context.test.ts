import { describe, expect, it } from 'vitest';

import { buildConversationContext, type ConversationContextMessage } from '../src/index.js';

const user = (content: string) => ({ role: 'user' as const, content });
const assistant = (content: string) => ({ role: 'assistant' as const, content });

function contents(messages: ConversationContextMessage[]): string[] {
  return messages.map((message) => message.content);
}

describe('buildConversationContext', () => {
  it('retains the most recent eligible turns and appends the current query', () => {
    const context = buildConversationContext(
      [
        { role: 'system', content: 'Do not expose this system instruction.' },
        user('The old request should fall out of the recent window.'),
        assistant('The old answer should fall out too.'),
        user('Keep this recent instruction.'),
        { role: 'tool', content: 'Tool output must never enter conversational context.' },
        assistant('Keep this recent answer.'),
      ],
      'Continue with the recent request.',
      { recentMessageCount: 2, maxRecalledMessages: 0 },
    );

    expect(contents(context.messages)).toEqual([
      'Keep this recent instruction.',
      'Keep this recent answer.',
      'Continue with the recent request.',
    ]);
    expect(
      context.messages.every((message) => message.role === 'user' || message.role === 'assistant'),
    ).toBe(true);
    expect(context.messages.at(-1)?.source).toBe('current-query');
  });

  it('recalls a relevant older user instruction by deterministic token overlap', () => {
    const context = buildConversationContext(
      [
        user('Always format revenue as currency and keep the amount precise.'),
        assistant('Understood.'),
        user('A later unrelated request about customers.'),
        assistant('The customer list is ready.'),
      ],
      'Please format the revenue column as currency.',
      { recentMessageCount: 2 },
    );

    expect(context.recalled).toHaveLength(1);
    expect(context.recalled[0]?.content).toContain('LOWER-TRUST CONVERSATIONAL CONTEXT');
    expect(context.recalled[0]?.overlapTokens).toEqual(
      expect.arrayContaining(['format', 'revenue', 'currency']),
    );
    expect(context.recalled[0]?.trust).toBe('lower-trust');
    expect(context.messages.some((message) => message.source === 'recalled')).toBe(true);
  });

  it('prioritizes an older correction over a stale instruction with more overlap', () => {
    const context = buildConversationContext(
      [
        user('Format revenue currency decimals and format revenue amount totals for every report.'),
        assistant('I will use those revenue formatting rules.'),
        user(
          'Actually, do not format revenue as currency; keep revenue as a plain number instead.',
        ),
        assistant('I will follow the correction.'),
        user('Please inspect the revenue values.'),
        assistant('I am inspecting them.'),
      ],
      'What should I do with revenue formatting?',
      { recentMessageCount: 2, maxRecalledMessages: 2 },
    );

    expect(context.recalled[0]?.content).toContain('Actually, do not format revenue as currency');
    expect(context.recalled[0]?.relevanceScore).toBeGreaterThan(
      context.recalled[1]?.relevanceScore ?? 0,
    );
  });

  it('enforces message and character caps even with oversized history and input', () => {
    const history = Array.from({ length: 100 }, (_, index) =>
      index % 2 === 0
        ? user(`Preference ${index}: always use the revenue format ${'x'.repeat(200)}`)
        : assistant(`Answer ${index}: ${'y'.repeat(200)}`),
    );
    const context = buildConversationContext(history, `revenue ${'z'.repeat(200)}`, {
      maxMessages: 5,
      recentMessageCount: 4,
      maxRecalledMessages: 10,
      maxCharacters: 180,
      maxMessageCharacters: 80,
    });

    expect(context.messages.length).toBeLessThanOrEqual(5);
    expect(context.characterCount).toBeLessThanOrEqual(180);
    expect(context.messages.every((message) => message.content.length <= 80)).toBe(true);
    expect(context.omittedMessageCount).toBeGreaterThan(0);
    expect(context.messages.at(-1)?.source).toBe('current-query');
  });
});
