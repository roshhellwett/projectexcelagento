"""Generate the M0 Excel fixtures.

This script is intentionally kept outside the TypeScript build. The generated files are
checked in so CI and future tests do not need Python or a fixture-generation dependency.
"""

from pathlib import Path

import xlsxwriter


ROOT = Path(__file__).parent


def new_workbook(name: str):
    return xlsxwriter.Workbook(ROOT / name)


def mixed_date_formats() -> None:
    workbook = new_workbook("01-mixed-date-formats.xlsx")
    sheet = workbook.add_worksheet("Orders")
    sheet.write_row(0, 0, ["Order", "Date", "Note"])
    sheet.write_row(1, 0, ["A-1", "01/02/2020", "Ambiguous"])
    sheet.write_row(2, 0, ["A-2", "31/12/2020", "Day first"])
    sheet.write_row(3, 0, ["A-3", "12/31/2020", "Month first"])
    sheet.write_row(4, 0, ["A-4", "2020-02-29", "ISO leap day"])
    workbook.close()


def merged_cells() -> None:
    workbook = new_workbook("02-merged-cells.xlsx")
    sheet = workbook.add_worksheet("Merged")
    title = workbook.add_format({"bold": True, "align": "center", "bg_color": "#D9EAF7"})
    sheet.merge_range("A1:D1", "Merged report title", title)
    sheet.write_row("A3", ["Region", "Order date", "Amount", "Status"])
    sheet.write_row("A4", ["West", "2024-01-15", 1200, "Open"])
    sheet.merge_range("A5:B5", "Merged label", title)
    sheet.write_row("C5", [800, "Closed"])
    workbook.close()


def blank_rows() -> None:
    workbook = new_workbook("03-blank-rows.xlsx")
    sheet = workbook.add_worksheet("Sparse")
    sheet.write_row(0, 0, ["ID", "Date", "Value"])
    sheet.write_row(1, 0, [1, 44927, 10])
    sheet.write_row(4, 0, [2, "2023-01-01", 20])
    sheet.write_row(7, 0, [3, "", 30])
    workbook.close()


def numbers_as_text() -> None:
    workbook = new_workbook("04-numbers-stored-as-text.xlsx")
    sheet = workbook.add_worksheet("Text numbers")
    text = workbook.add_format({"num_format": "@"})
    sheet.write_row(0, 0, ["ID", "Amount", "Potential serial"])
    sheet.write_string(1, 0, "001", text)
    sheet.write_string(1, 1, "1200.50", text)
    sheet.write_string(1, 2, "44927", text)
    sheet.write_string(2, 0, "002", text)
    sheet.write_string(2, 1, "35", text)
    sheet.write_string(2, 2, "2024-01-01", text)
    workbook.close()


def multiple_sheets() -> None:
    workbook = new_workbook("05-multiple-sheets.xlsx")
    orders = workbook.add_worksheet("Orders")
    orders.write_row(0, 0, ["Order", "Date", "Amount"])
    orders.write_row(1, 0, ["A-1", "2024-02-29", 20])
    customers = workbook.add_worksheet("Customers")
    customers.write_row(0, 0, ["Customer", "Joined"])
    customers.write_row(1, 0, ["Ada", "31/01/2024"])
    notes = workbook.add_worksheet("Notes")
    notes.write("A1", "This sheet intentionally has no table.")
    workbook.close()


def chart() -> None:
    workbook = new_workbook("06-chart.xlsx")
    sheet = workbook.add_worksheet("Revenue")
    sheet.write_row(0, 0, ["Month", "Revenue"])
    values = [("Jan", 120), ("Feb", 160), ("Mar", 145), ("Apr", 210)]
    for row, (month, revenue) in enumerate(values, start=1):
        sheet.write(row, 0, month)
        sheet.write(row, 1, revenue)
    chart = workbook.add_chart({"type": "column"})
    chart.add_series(
        {
            "name": "Revenue",
            "categories": "=Revenue!$A$2:$A$5",
            "values": "=Revenue!$B$2:$B$5",
        }
    )
    chart.set_title({"name": "Monthly revenue"})
    chart.set_legend({"none": True})
    sheet.insert_chart("D2", chart)
    workbook.close()


def conditional_formatting() -> None:
    workbook = new_workbook("07-conditional-formatting.xlsx")
    sheet = workbook.add_worksheet("Scores")
    sheet.write_row(0, 0, ["Student", "Score", "Date"])
    rows = [("Ada", 98, "2024-02-29"), ("Lin", 72, "31/01/2024"), ("Sam", 45, "01/02/2024")]
    for row, values in enumerate(rows, start=1):
        sheet.write_row(row, 0, values)
    sheet.conditional_format("B2:B4", {"type": "3_color_scale"})
    workbook.close()


def leap_days() -> None:
    workbook = new_workbook("08-leap-days.xlsx")
    sheet = workbook.add_worksheet("Dates")
    sheet.write_row(0, 0, ["Label", "Date"])
    sheet.write_row(1, 0, ["Valid leap day", "2020-02-29"])
    sheet.write_row(2, 0, ["Non-leap year", "29/02/2019"])
    sheet.write_row(3, 0, ["Excel serial leap day", 43890])
    workbook.close()


def formulas_and_empty_cells() -> None:
    workbook = new_workbook("09-formulas-and-empty-cells.xlsx")
    sheet = workbook.add_worksheet("Calculated")
    sheet.write_row(0, 0, ["Date", "Amount", "Total"])
    sheet.write(1, 0, 43890)
    sheet.write(1, 1, 10)
    sheet.write_formula(1, 2, "=B2*2", None, 20)
    sheet.write_blank(2, 0, None)
    sheet.write(2, 1, 25)
    sheet.write_formula(2, 2, "=B3*2", None, 50)
    workbook.close()


def combined_messy() -> None:
    workbook = new_workbook("10-combined-messy-orders.xlsx")
    sheet = workbook.add_worksheet("Orders")
    header = workbook.add_format({"bold": True, "bg_color": "#E2F0D9"})
    sheet.merge_range("A1:D1", "Imported orders", header)
    sheet.write_row(1, 0, ["Order", "Date", "Amount", "Region"], header)
    sheet.write_row(2, 0, ["A-1", 43890, "1200", "West"])
    sheet.write_row(3, 0, ["A-2", "01/02/2020", "950.25", "East"])
    sheet.write_row(4, 0, ["A-3", "31/12/2020", 700, "West"])
    sheet.write_row(6, 0, ["A-4", "2024-02-29", 825, "North"])
    sheet.conditional_format("C3:C7", {"type": "data_bar", "bar_color": "#63C384"})
    workbook.close()


def main() -> None:
    mixed_date_formats()
    merged_cells()
    blank_rows()
    numbers_as_text()
    multiple_sheets()
    chart()
    conditional_formatting()
    leap_days()
    formulas_and_empty_cells()
    combined_messy()


if __name__ == "__main__":
    main()
