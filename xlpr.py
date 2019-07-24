import argparse
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Border, Side, Alignment, PatternFill, Font
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule, Rule

parser = argparse.ArgumentParser(description='Automate creation of scoring Excel worksheets.')
parser.add_argument('name', help='What to call the file')
parser.add_argument('num_questions', type=int, help='How many question columns to generate')
parser.add_argument('num_participants', type=int, help='How many question columns to generate')
parser.add_argument('--range-high', dest='rangeHigh', type=int, help='If there is a range for all questions, what is the top of that range for conditional formatting?')
args = parser.parse_args()

def do_borders(ws, x):
    for col in range(1, x):
        ws.cell(column=col, row=1).border = Border(bottom=Side(border_style="thin", color="000000"))

def fill_sheet(ws, rangeHigh, rangeLow=1):
    # Add headers
    ws['A1'] = 'Sub ID'
    ws['B1'] = 'Date'
    ws['C1'] = 'Entered By'
    ws.column_dimensions["A"].width = "10"
    ws.column_dimensions["B"].width = "10"
    ws.column_dimensions["C"].width = "15"
    for col in range(4, 4 + args.num_questions):
        ws.cell(column=col, row=1, value="Q{0}".format(col-3))
        ws.column_dimensions[get_column_letter(col)].width = "8"

    ws.cell(column=4 + args.num_questions, row=1, value="Notes")

    # Add subject numbers
    for row in range(2, 2 + args.num_participants):
        ws.cell(column=1, row=row, value=(row-1))

    # Freeze row 1, columns 1-3 - basically this freezes everything "above and to the left" of the given cell
    ws.freeze_panes = "D2"

    # TODO: Maybe notes should be 4 and maybe frozen?
    do_borders(ws, 5 + args.num_questions)
    
    last_cell = ws.cell(column=4 + args.num_questions, row=2 + args.num_participants)

    # Conditional formatting to do blanks as light gray
    blankFill = PatternFill(start_color='EEEEEE', end_color='EEEEEE', fill_type='solid')
    dxf = DifferentialStyle(fill=blankFill)
    blank = Rule(type="expression", formula = ["ISBLANK(D2)"], dxf=dxf, stopIfTrue=True)
    ws.conditional_formatting.add('D2:{0}'.format(last_cell.coordinate), blank)
                              
    # Conditional formatting to highlight numbers out of range
    if rangeHigh:
        yellowFill = PatternFill(start_color='EEEE66', end_color='EEEE66', fill_type='solid')
        rule = CellIsRule(operator="notBetween", formula=[str(rangeLow),str(rangeHigh)], fill=yellowFill)
        ws.conditional_formatting.add('D2:{0}'.format(last_cell.coordinate), rule)

    

def compare_sheet(ws):
    # Add headers
    ws['A1'] = 'Sub ID'
    ws['C1'] = 'Bad'
    for col in range(4, 4 + args.num_questions):
        ws.cell(column=col, row=1, value="Q{0}".format(col-3))
    for row in range(2, 2 + args.num_participants):
        rater1 = ws.cell(column=2,row=row)
        rater1.value = "=Entry1!C{0}".format(row)
        rater2 = ws.cell(column=3,row=row)
        rater2.value = "=Entry2!C{0}".format(row)
        for col in range(4, 4 + args.num_questions):
            cell = ws.cell(column=col,row=row)
            cell.value = "=IF(Entry1!{0}=Entry2!{0},Entry2!{0},CONCATENATE(Entry1!{0},\" vs. \",Entry2!{0}))".format(cell.coordinate)

    last_cell = ws.cell(column=4 + args.num_questions, row=2 + args.num_participants)

    # Conditional formatting to highlight mismatches
    redFill = PatternFill(start_color='EE6666', end_color='EE6666', fill_type='solid')
    dxf = DifferentialStyle(fill=redFill)
    rule = Rule(type="containsText", operator="containsText", text="vs.", dxf=dxf)
    ws.conditional_formatting.add('D2:{0}'.format(last_cell.coordinate), rule)

wb = Workbook()

wb.remove(wb.active)
fill_sheet(wb.create_sheet("Entry1"), args.rangeHigh)
fill_sheet(wb.create_sheet("Entry2"), args.rangeHigh)
compare_sheet(wb.create_sheet("Final_Comparison"))

# Save the file
wb.save(args.name + ".xlsx")

