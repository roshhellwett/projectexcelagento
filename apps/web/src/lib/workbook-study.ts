import { maxColumnCount, type Workbook } from '@excel-agent/engine';

/**
 * Performs a comprehensive automated study and structural audit of an uploaded Excel workbook.
 * Produces a clear, high-level analytical briefing for the user upon document ingestion.
 */
export function generateWorkbookStudy(workbook: Workbook, fileName: string): string {
  const totalSheets = workbook.sheets.length;
  const totalRows = workbook.sheets.reduce((sum, s) => sum + s.rows.length, 0);

  const sheetSummaries = workbook.sheets.map((sheet) => {
    const rowCount = sheet.rows.length;
    const colCount = maxColumnCount(sheet.rows);

    const firstRowValues = (sheet.rows[0] ?? [])
      .map((c) => String(c?.value ?? '').trim())
      .filter(Boolean);

    // Detect single-column code/script sheet
    const isScriptOrLog =
      colCount === 1 &&
      rowCount > 3 &&
      sheet.rows.slice(0, 15).some((r) => {
        const val = String(r[0]?.value ?? '');
        return (
          val.startsWith('#!') ||
          val.includes('import ') ||
          val.includes('python ') ||
          val.includes('def ') ||
          val.includes('export ') ||
          val.includes('const ')
        );
      });

    // Detect structured template with headers but 0 or few data rows
    const isStructuredTemplate = firstRowValues.length >= 3 && rowCount <= 3;

    // Detect normal tabular data
    let classification = 'Tabular Dataset';
    let details = '';

    if (isScriptOrLog) {
      classification = 'Raw Source Script / Unstructured Data';
      details = `Single-column script (${rowCount} lines). Contains unparsed raw entities/records waiting to be structured.`;
    } else if (isStructuredTemplate) {
      classification = 'Structured Target Schema';
      details = `Predefined table with ${firstRowValues.length} columns: \`[${firstRowValues.slice(0, 8).join(', ')}${firstRowValues.length > 8 ? '...' : ''}]\`. Currently 0 data rows (ready to be populated).`;
    } else {
      classification = 'Tabular Dataset';
      const numColsCount = (sheet.rows[1] ?? []).filter((c) => typeof c?.value === 'number').length;
      details = `${rowCount} rows × ${colCount} columns. Headers: \`[${firstRowValues.slice(0, 7).join(', ')}${firstRowValues.length > 7 ? '...' : ''}]\`. Detected ~${numColsCount} numeric column(s).`;
    }

    return {
      sheet,
      name: sheet.name,
      rowCount,
      colCount,
      classification,
      details,
      isScriptOrLog,
      isStructuredTemplate,
      firstRowValues,
    };
  });

  const scriptSheet = sheetSummaries.find((s) => s.isScriptOrLog);
  const templateSheet = sheetSummaries.find((s) => s.isStructuredTemplate);

  const lines: string[] = [
    `### 📁 Automated Ingestion & Deep Study: **"${fileName}"**`,
    '',
    `I have analyzed your workbook across **${totalSheets} sheet(s)** with **${totalRows.toLocaleString()} total rows**. Here is the structural breakdown:`,
    '',
  ];

  sheetSummaries.forEach((s, idx) => {
    lines.push(`**${idx + 1}. \`${s.name}\`** — *${s.classification}*`);
    lines.push(`• ${s.details}`);
    lines.push('');
  });

  lines.push('---');
  lines.push('#### 💡 Next Recommended Steps:');

  if (scriptSheet && templateSheet) {
    lines.push(
      `• **Populate Structured Table:** Detected raw records in **\`${scriptSheet.name}\`** and a structured schema in **\`${templateSheet.name}\`**. Say **"now fill the data"** or **"data daal isme"** to extract and populate leads automatically.`,
    );
    lines.push(
      `• **Analyze Code Pipeline:** Ask **"what does ${scriptSheet.name} do?"** to understand the scraping parameters or API dependencies.`,
    );
  } else {
    lines.push(
      '• **Clean & Normalize:** Ask to **"clean this sheet"** to standardize formatting, trim whitespace, and fix dates.',
    );
    lines.push(
      '• **Deduplication:** Ask to **"remove duplicate rows"** to inspect and remove repeating entries.',
    );
    lines.push(
      '• **Financial & Numerical Totals:** Ask for **"total sum of [column]"** or **"top 5 highest amounts"** for instant verified calculations.',
    );
  }

  lines.push('');
  lines.push('*Type your instruction below to begin!*');

  return lines.join('\n');
}
