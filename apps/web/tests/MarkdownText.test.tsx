// @vitest-environment jsdom
import React from 'react';
import { describe, it, expect } from 'vitest';
import { render, screen } from '@testing-library/react';
import { MarkdownText } from '../src/components/MarkdownText.js';

describe('MarkdownText', () => {
  it('renders headings, paragraphs, and inline styles', () => {
    const md = '### Financial Summary\nThis is **bold** and `code`.';
    const { container } = render(<MarkdownText text={md} />);
    expect(screen.getByRole('heading', { level: 5 })).toHaveTextContent('Financial Summary');
    expect(container.querySelector('strong')).toHaveTextContent('bold');
    expect(container.querySelector('code')).toHaveTextContent('code');
  });

  it('renders standard markdown tables into structured HTML table elements', () => {
    const tableMd = `
| Year | Revenue | Operating Income | Margin |
|---|---|---|---|
| FY09 | 30,990 | 8,231 | 26.6% |
| FY10 | 35,119 | 8,413 | 24.0% |
`;
    const { container } = render(<MarkdownText text={tableMd} />);
    const table = container.querySelector('table.md-table');
    expect(table).not.toBeNull();

    const headers = container.querySelectorAll('th');
    expect(headers).toHaveLength(4);
    expect(headers[0]).toHaveTextContent('Year');
    expect(headers[1]).toHaveTextContent('Revenue');

    const cells = container.querySelectorAll('td');
    expect(cells).toHaveLength(8);
    expect(cells[0]).toHaveTextContent('FY09');
    expect(cells[1]).toHaveTextContent('30,990');
  });

  it('parses collapsed/joined table rows with double pipes into structured table', () => {
    // This matches the exact format seen in LLM streaming/collapsed output
    const collapsedMd =
      '| Year | Revenue | Operating Income ||---|---|---|| FY09 | 30,990 | 8,231 || FY10 | 35,119 | 8,413 |';
    const { container } = render(<MarkdownText text={collapsedMd} />);
    const table = container.querySelector('table.md-table');
    expect(table).not.toBeNull();

    const rows = container.querySelectorAll('tbody tr');
    expect(rows.length).toBeGreaterThanOrEqual(2);
  });
});
