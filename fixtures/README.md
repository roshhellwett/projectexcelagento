# Excel fixtures

These ten `.xlsx` files are small, intentionally messy inputs for the upcoming parser and fidelity scan:

| File                               | Coverage                                                           |
| ---------------------------------- | ------------------------------------------------------------------ |
| `01-mixed-date-formats.xlsx`       | ISO, day-first, month-first, and ambiguous date text               |
| `02-merged-cells.xlsx`             | Merged title and label cells                                       |
| `03-blank-rows.xlsx`               | Blank rows inside a table                                          |
| `04-numbers-stored-as-text.xlsx`   | Numeric IDs, amounts, and serial-looking text stored as strings    |
| `05-multiple-sheets.xlsx`          | Three worksheets                                                   |
| `06-chart.xlsx`                    | Embedded column chart                                              |
| `07-conditional-formatting.xlsx`   | Three-color-scale conditional formatting                           |
| `08-leap-days.xlsx`                | Valid and invalid leap-day text plus an Excel serial               |
| `09-formulas-and-empty-cells.xlsx` | Formula cells and blank cells                                      |
| `10-combined-messy-orders.xlsx`    | Merged cells, blank rows, mixed values, and conditional formatting |

To regenerate them locally, install `XlsxWriter` in a temporary Python virtual environment and run:

```bash
python3 -m venv /tmp/excel-agent-fixtures-venv
/tmp/excel-agent-fixtures-venv/bin/pip install XlsxWriter
/tmp/excel-agent-fixtures-venv/bin/python fixtures/generate_fixtures.py
```
