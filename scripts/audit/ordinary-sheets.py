"""Builds ordinary workbooks: the kind people really keep, each with one everyday feature.

Usage: python3 ordinary-sheets.py OUTPUT_DIRECTORY
"""
import sys, os
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.workbook.defined_name import DefinedName

out = sys.argv[1]
os.makedirs(out, exist_ok=True)
ROWS = [("North", "Widget", 12, 4.5), ("South", "Gadget", 7, 12.0), ("East", "Widget", 20, 4.5),
        ("West", "Gizmo", 3, 30.25), ("North", "Gizmo", 9, 30.25), ("South", "Widget", 15, 4.5)]

def base(title_rows=0):
    book = Workbook(); sheet = book.active; sheet.title = "Sales"
    top = title_rows + 1
    for col, name in enumerate(["Region", "Item", "Quantity", "Price", "Total"], 1):
        cell = sheet.cell(row=top, column=col, value=name); cell.font = Font(bold=True)
    for index, (region, item, quantity, price) in enumerate(ROWS):
        row = top + 1 + index
        sheet.cell(row=row, column=1, value=region); sheet.cell(row=row, column=2, value=item)
        sheet.cell(row=row, column=3, value=quantity); sheet.cell(row=row, column=4, value=price)
        sheet.cell(row=row, column=5, value=f"=C{row}*D{row}")
    last = top + len(ROWS)
    sheet.cell(row=last + 1, column=1, value="Total")
    sheet.cell(row=last + 1, column=5, value=f"=SUM(E{top + 1}:E{last})")
    return book, sheet, top, last

def save(name, book):
    book.save(os.path.join(out, name + ".xlsx"))

book, sheet, top, last = base(); save("01-plain", book)
book, sheet, top, last = base(); sheet.auto_filter.ref = f"A{top}:E{last}"; save("02-filter", book)
book, sheet, top, last = base(1); sheet["A1"] = "Quarterly sales"; sheet.merge_cells("A1:E1"); sheet["A1"].alignment = Alignment(horizontal="center"); save("03-merged-title", book)
book, sheet, top, last = base()
table = Table(displayName="SalesTable", ref=f"A{top}:E{last}"); table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
sheet.add_table(table); save("04-table", book)
book, sheet, top, last = base()
sheet.conditional_formatting.add(f"C{top + 1}:C{last}", CellIsRule(operator="greaterThan", formula=["10"], fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")))
save("05-highlight-rule", book)
book, sheet, top, last = base()
sheet.conditional_formatting.add(f"E{top + 1}:E{last}", ColorScaleRule(start_type="min", start_color="F8696B", end_type="max", end_color="63BE7B"))
save("06-colour-scale", book)
book, sheet, top, last = base(); sheet["G2"] = "Supplier site"; sheet["G2"].hyperlink = "https://example.com/supplier"; save("07-hyperlink", book)
book, sheet, top, last = base()
rule = DataValidation(type="list", formula1='"North,South,East,West"', allow_blank=True); sheet.add_data_validation(rule); rule.add(f"A{top + 1}:A{last}")
save("08-dropdown", book)
book, sheet, top, last = base(); sheet.freeze_panes = "A2"; sheet.column_dimensions["A"].width = 18
for row in range(top + 1, last + 1): sheet.cell(row=row, column=4).number_format = "#,##0.00"
save("09-frozen-formatted", book)
book, sheet, top, last = base()
chart = BarChart(); chart.add_data(Reference(sheet, min_col=5, min_row=top, max_row=last), titles_from_data=True); chart.set_categories(Reference(sheet, min_col=1, min_row=top + 1, max_row=last)); sheet.add_chart(chart, "H2")
save("10-chart", book)
book, sheet, top, last = base()
rates = book.create_sheet("Rates"); rates["A1"] = "Tax"; rates["B1"] = 0.09
book.defined_names["TaxRate"] = DefinedName("TaxRate", attr_text="Rates!$B$1")
sheet["F1"] = "With tax"
for row in range(top + 1, last + 1): sheet.cell(row=row, column=6, value=f"=E{row}*(1+TaxRate)")
save("11-names-two-sheets", book)
book, sheet, top, last = base(); sheet["C2"].comment = Comment("Check this quantity", "Reviewer"); save("12-note", book)
book, sheet, top, last = base(1)
sheet["A1"] = "Quarterly sales"; sheet.merge_cells("A1:E1"); sheet.auto_filter.ref = f"A{top}:E{last}"; sheet.freeze_panes = f"A{top + 1}"
sheet.conditional_formatting.add(f"C{top + 1}:C{last}", CellIsRule(operator="greaterThan", formula=["10"], fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")))
rule = DataValidation(type="list", formula1='"North,South,East,West"', allow_blank=True); sheet.add_data_validation(rule); rule.add(f"A{top + 1}:A{last}")
sheet["G3"] = "Supplier site"; sheet["G3"].hyperlink = "https://example.com/supplier"; sheet["C3"].comment = Comment("Check this quantity", "Reviewer")
save("13-typical-report", book)
# Manual page breaks after the fourth and seventh rows and after the third column.
from openpyxl.worksheet.pagebreak import Break
book, sheet, top, last = base()
for row in (4, 7): sheet.row_breaks.append(Break(id=row))
sheet.col_breaks.append(Break(id=3))
save("16-page-breaks", book)
# A filter as Excel leaves it after "sort largest first" on Quantity, with the
# dropdown button of the first column hidden.
from openpyxl.worksheet.filters import FilterColumn, SortState, SortCondition
book, sheet, top, last = base()
sheet.auto_filter.ref = f"A{top}:E{last}"
sheet.auto_filter.filterColumn.append(FilterColumn(colId=0, hiddenButton=True))
sheet.auto_filter.sortState = SortState(ref=f"A{top + 1}:E{last}", sortCondition=[SortCondition(ref=f"C{top + 1}:C{last}", descending=True)])
save("17-sorted-filter", book)
# One protected sheet beside an ordinary one.
book, sheet, top, last = base()
sheet.protection.sheet = True
notes = book.create_sheet("Notes"); notes["A1"] = "Reviewed"; notes["B1"] = f"=Sales!E{last + 1}"
save("18-protected-sheet", book)
# A table, and a formula beside it that names the table.
book, sheet, top, last = base()
table = Table(displayName="Orders", ref=f"A{top}:E{last}"); table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
sheet.add_table(table); sheet["H1"] = "=SUM(Orders[Total])"
save("19-table-formula", book)
print("built 13 workbooks in", out)
# Sheet names in another script, read across sheets without quotes, as Excel writes them.
book, sheet, top, last = base(); sheet.title = "数据"
sheet["G2"] = "供应商"; sheet["G2"].hyperlink = "https://example.com/supplier"
summary = book.create_sheet("汇总")
summary["A1"] = "合计"; summary["B1"] = f"=SUM(数据!E{top + 1}:E{last})"; summary["B2"] = "=数据!C2*2"
save("15-other-language", book)
