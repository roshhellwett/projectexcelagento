import { describe, expect, it } from 'vitest';

import { createOperationRegistry } from '@excel-agent/engine';

import { ToolCatalogError, buildToolCatalog, describeTools } from '../src/index.js';

describe('tool catalog', () => {
  it('derives one tool per engine operation with a JSON schema', () => {
    const registry = createOperationRegistry();
    const catalog = buildToolCatalog(registry);

    expect(catalog.map((tool) => tool.name)).toEqual(registry.names);
    for (const tool of catalog) {
      expect(tool.jsonSchema.type).toBe('object');
      expect(tool.jsonSchema.properties).toBeTypeOf('object');
      expect(tool.description.length).toBeGreaterThan(0);
    }
  });

  it('documents every engine operation the model is allowed to call', () => {
    const catalog = buildToolCatalog(createOperationRegistry());
    const byName = new Map(catalog.map((tool) => [tool.name, tool]));

    // The five business operations, end to end: live in the registry, survive catalog construction
    // (which throws on an undocumented name), and reach the model with a usable contract.
    for (const name of [
      'aggregate_column',
      'group_and_summarize',
      'join_sheets',
      'fill_series',
      'categorize_column',
    ]) {
      const tool = byName.get(name);
      expect(tool, `${name} is missing from the tool catalog`).toBeDefined();
      expect(tool?.description.length ?? 0).toBeGreaterThan(40);
      expect(tool?.example).toMatchObject({ sheet: expect.any(String) });
      expect((tool?.jsonSchema.properties as Record<string, unknown>)?.sheet).toBeDefined();
      expect(describeTools(catalog)).toContain(`"name":"${name}"`);
    }
  });

  it('throws when the engine exposes an undocumented operation', () => {
    expect(() =>
      buildToolCatalog({
        names: ['mystery_operation'],
        getSchema: () => undefined,
      } as unknown as Parameters<typeof buildToolCatalog>[0]),
    ).toThrow(ToolCatalogError);
  });

  it('renders a model-friendly contract that names every operation', () => {
    const catalog = buildToolCatalog(createOperationRegistry());
    const rendered = describeTools(catalog);

    for (const tool of catalog) {
      expect(rendered).toContain(tool.name);
    }
    expect(rendered).toContain('"name":"format_dates"');
  });
});
