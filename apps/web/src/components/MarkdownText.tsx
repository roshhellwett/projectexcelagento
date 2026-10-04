import React from 'react';

interface MarkdownTextProps {
  text: string;
  className?: string;
}

/** Lightweight structured renderer for assistant replies (headings, lists, bold, inline code, fenced code). */
export const MarkdownText: React.FC<MarkdownTextProps> = ({ text, className }) => {
  const blocks = React.useMemo(() => parseBlocks(text), [text]);
  return (
    <div className={className ? `md-content ${className}` : 'md-content'}>
      {blocks.map((block, i) => (
        <BlockView key={i} block={block} />
      ))}
    </div>
  );
};

type Block =
  | { type: 'heading'; level: number; text: string }
  | { type: 'code'; code: string }
  | { type: 'list'; ordered: boolean; items: string[] }
  | { type: 'quote'; text: string }
  | {
      type: 'table';
      headers: string[];
      alignments: ('left' | 'center' | 'right')[];
      rows: string[][];
    }
  | { type: 'paragraph'; text: string };

function isDelimiterRow(line: string): boolean {
  const trimmed = line.trim();
  if (!trimmed.includes('-')) return false;
  return /^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)+\|?\s*$/.test(trimmed);
}

function extractCells(line: string): string[] {
  let s = line.trim();
  if (s.startsWith('|')) s = s.slice(1);
  if (s.endsWith('|')) s = s.slice(0, -1);
  return s.split('|').map((c) => c.trim());
}

function parseAlignments(delimiterLine: string): ('left' | 'center' | 'right')[] {
  return extractCells(delimiterLine).map((cell) => {
    const trimmed = cell.trim();
    const startColon = trimmed.startsWith(':');
    const endColon = trimmed.endsWith(':');
    if (startColon && endColon) return 'center';
    if (endColon) return 'right';
    return 'left';
  });
}

function normalizeMarkdownText(raw: string): string[] {
  let preprocessed = raw
    .replace(/\|\s*\|\s*---/g, '|\n|---')
    .replace(/\|\s*\|\s*/g, '|\n|')
    .replace(/\|\s*(\r?\n)+\s*\|/g, '|\n|')
    .replace(/\|\s+\|\s*---/g, '|\n|---');
  return preprocessed.split(/\r?\n/);
}

function parseBlocks(text: string): Block[] {
  const lines = normalizeMarkdownText(text);
  const blocks: Block[] = [];
  let i = 0;
  while (i < lines.length) {
    const line = lines[i] ?? '';
    if (/^```/.test(line.trim())) {
      const codeLines: string[] = [];
      i += 1;
      while (i < lines.length && !/^```/.test(lines[i]!.trim())) {
        codeLines.push(lines[i] ?? '');
        i += 1;
      }
      i += 1;
      blocks.push({ type: 'code', code: codeLines.join('\n') });
      continue;
    }
    if (line.trim() === '') {
      i += 1;
      continue;
    }

    // Markdown Table detection
    if (
      line.includes('|') &&
      i + 1 < lines.length &&
      isDelimiterRow(lines[i + 1] ?? '')
    ) {
      const headers = extractCells(line);
      const alignments = parseAlignments(lines[i + 1] ?? '');
      i += 2;
      const rows: string[][] = [];
      while (i < lines.length && (lines[i] ?? '').includes('|') && (lines[i] ?? '').trim() !== '') {
        const rowCells = extractCells(lines[i] ?? '');
        rows.push(rowCells);
        i += 1;
      }
      blocks.push({ type: 'table', headers, alignments, rows });
      continue;
    }

    const heading = /^(#{1,3})\s+(.*)$/.exec(line);
    if (heading) {
      blocks.push({ type: 'heading', level: heading[1]!.length, text: heading[2]! });
      i += 1;
      continue;
    }
    if (/^\s*>\s?/.test(line)) {
      blocks.push({ type: 'quote', text: line.replace(/^\s*>\s?/, '') });
      i += 1;
      continue;
    }
    if (/^\s*[-*•]\s+/.test(line)) {
      const items: string[] = [];
      while (i < lines.length && /^\s*[-*•]\s+/.test(lines[i]!)) {
        items.push(lines[i]!.replace(/^\s*[-*•]\s+/, ''));
        i += 1;
      }
      blocks.push({ type: 'list', ordered: false, items });
      continue;
    }
    if (/^\s*\d+[.)]\s+/.test(line)) {
      const items: string[] = [];
      while (i < lines.length && /^\s*\d+[.)]\s+/.test(lines[i]!)) {
        items.push(lines[i]!.replace(/^\s*\d+[.)]\s+/, ''));
        i += 1;
      }
      blocks.push({ type: 'list', ordered: true, items });
      continue;
    }

    // Paragraph: gather consecutive non-empty, non-special lines
    const paraLines: string[] = [];
    while (
      i < lines.length &&
      lines[i]!.trim() !== '' &&
      !/^```/.test(lines[i]!.trim()) &&
      !/^(#{1,3})\s+/.test(lines[i]!) &&
      !/^\s*[-*•]\s+/.test(lines[i]!) &&
      !/^\s*\d+[.)]\s+/.test(lines[i]!) &&
      !/^\s*>\s?/.test(lines[i]!) &&
      !(lines[i]!.includes('|') && i + 1 < lines.length && isDelimiterRow(lines[i + 1] ?? ''))
    ) {
      paraLines.push(lines[i] ?? '');
      i += 1;
    }
    if (paraLines.length > 0) {
      blocks.push({ type: 'paragraph', text: paraLines.join(' ') });
    } else {
      i += 1;
    }
  }
  return blocks;
}

const BlockView: React.FC<{ block: Block }> = ({ block }) => {
  switch (block.type) {
    case 'heading': {
      const Tag = block.level === 1 ? 'h3' : block.level === 2 ? 'h4' : 'h5';
      return <Tag className="md-heading">{renderInline(block.text)}</Tag>;
    }
    case 'code':
      return (
        <pre className="md-code">
          <code>{block.code}</code>
        </pre>
      );
    case 'list': {
      const Tag = block.ordered ? 'ol' : 'ul';
      return (
        <Tag className="md-list">
          {block.items.map((item, idx) => (
            <li key={idx}>{renderInline(item)}</li>
          ))}
        </Tag>
      );
    }
    case 'quote':
      return <blockquote className="md-quote">{renderInline(block.text)}</blockquote>;
    case 'table':
      return (
        <div className="md-table-wrapper" tabIndex={0}>
          <table className="md-table">
            <thead>
              <tr>
                {block.headers.map((hdr, hIdx) => (
                  <th
                    key={hIdx}
                    style={{ textAlign: block.alignments[hIdx] ?? 'left' }}
                  >
                    {renderInline(hdr)}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {block.rows.map((row, rIdx) => (
                <tr key={rIdx}>
                  {row.map((cell, cIdx) => (
                    <td
                      key={cIdx}
                      style={{ textAlign: block.alignments[cIdx] ?? 'left' }}
                    >
                      {renderInline(cell)}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      );
    case 'paragraph':
      return <p className="md-paragraph">{renderInline(block.text)}</p>;
  }
};

function renderInline(text: string): React.ReactNode {
  const parts = text.split(/(\*\*[^*]+\*\*|`[^`]+`|\*[^*]+\*)/g);
  return parts.map((part, idx) => {
    if (part.startsWith('**') && part.endsWith('**') && part.length > 4) {
      return <strong key={idx}>{part.slice(2, -2)}</strong>;
    }
    if (part.startsWith('`') && part.endsWith('`') && part.length > 2) {
      return (
        <code key={idx} className="md-inline-code">
          {part.slice(1, -1)}
        </code>
      );
    }
    if (part.startsWith('*') && part.endsWith('*') && part.length > 2) {
      return <em key={idx}>{part.slice(1, -1)}</em>;
    }
    return <React.Fragment key={idx}>{part}</React.Fragment>;
  });
}
