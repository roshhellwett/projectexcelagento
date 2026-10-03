import { createCell, type Workbook } from '@excel-agent/engine';

export const EVAL_SHEET = 'Orders';

/**
 * Canonical fixture shared by every eval case. It intentionally contains a
 * duplicate row, mixed date formats, padded text, and a numeric column so the
 * golden prompts exercise real cleaning paths.
 */
export function evalWorkbook(): Workbook {
  return {
    sheets: [
      {
        name: EVAL_SHEET,
        rows: [
          [
            createCell('Order ID'),
            createCell('Customer'),
            createCell('Order Date'),
            createCell('Amount'),
            createCell('Status'),
          ],
          [
            createCell('ORD-1001'),
            createCell('  acme corp  '),
            createCell('2026-03-01'),
            createCell(1420.5),
            createCell('Completed'),
          ],
          [
            createCell('ORD-1002'),
            createCell('bmc ltd'),
            createCell('03/15/2026'),
            createCell(840),
            createCell('Pending'),
          ],
          [
            createCell('ORD-1003'),
            createCell('Acme Corp'),
            createCell('2026/04/05'),
            createCell(2150),
            createCell('Completed'),
          ],
          [
            createCell('ORD-1002'),
            createCell('bmc ltd'),
            createCell('03/15/2026'),
            createCell(840),
            createCell('Pending'),
          ],
          [
            createCell('ORD-1004'),
            createCell('apex holdings'),
            createCell('18/05/2026'),
            createCell(5400),
            createCell('Completed'),
          ],
        ],
      },
    ],
  };
}
