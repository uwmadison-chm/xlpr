import argparse
import xlrd
import os
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Border, Side, Alignment, PatternFill, Font
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule, Rule

parser = argparse.ArgumentParser(description='Automate creation of scoring Excel worksheets.')
subparsers = parser.add_subparsers(dest='subcommand', title='subcommands', description='valid subcommands', required=True)

manual = subparsers.add_parser('manual')
manual.add_argument('name', help='What to call the file')
manual.add_argument('num_questions', type=int, help='How many question columns to generate')
manual.add_argument('num_participants', type=int, help='How many question columns to generate')
manual.add_argument('--range-high', dest='range_high', type=int, help='If there is a range for all questions, what is the top of that range for conditional formatting?')

auto = subparsers.add_parser('auto')
auto.add_argument('name', help='Input excel file that specifies what questionairres to create')
auto.add_argument('num_participants', type=int, help='How many question columns to generate')
auto.add_argument('output_path', help='Where to create the files')

args = parser.parse_args()

def do_borders(ws, x):
    for col in range(1, x):
        ws.cell(column=col, row=1).border = Border(bottom=Side(border_style="thin", color="000000"))

def fill_sheet(ws, num_questions, num_participants, range_high, range_low=1, copyRow2FromSheet1=False):
    # Add headers
    ws['A1'] = 'Sub ID'
    ws['B1'] = 'Date'
    ws['C1'] = 'Entered By'
    ws.column_dimensions["A"].width = "10"
    ws.column_dimensions["B"].width = "10"
    ws.column_dimensions["C"].width = "15"
    for col in range(4, 4 + num_questions):
        ws.cell(column=col, row=1, value="Q{0}".format(col-3))
        ws.column_dimensions[get_column_letter(col)].width = "8"
        if copyRow2FromSheet1:
            question_metadata = ws.cell(column=col,row=2)
            question_metadata.value = '=Entry1!{0}'.format(question_metadata.coordinate)

    ws.cell(column=4 + num_questions, row=1, value="Notes")

    # Add subject numbers
    for row in range(3, 3 + num_participants):
        ws.cell(column=1, row=row, value=(row-2))

    # Freeze row 1-2, columns 1-3 - basically this freezes everything "above and to the left" of the given cell
    ws.freeze_panes = "D3"

    do_borders(ws, 5 + num_questions)
    
    last_cell = ws.cell(column=3 + num_questions, row=2 + num_participants)

    # Conditional formatting to do blanks as light gray
    blankFill = PatternFill(start_color='EEEEEE', end_color='EEEEEE', fill_type='solid')
    dxf = DifferentialStyle(fill=blankFill)
    blank = Rule(type="expression", formula = ["ISBLANK(D3)"], dxf=dxf, stopIfTrue=True)
    ws.conditional_formatting.add('D3:{0}'.format(last_cell.coordinate), blank)
                              
    # Conditional formatting to highlight numbers out of range
    if range_high:
        yellowFill = PatternFill(start_color='EEEE66', end_color='EEEE66', fill_type='solid')
        rule = CellIsRule(operator="notBetween", formula=[str(range_low),str(range_high)], fill=yellowFill)
        ws.conditional_formatting.add('D3:{0}'.format(last_cell.coordinate), rule)

    

def compare_sheet(ws, num_questions, num_participants):
    # Add headers
    ws['A1'] = 'Sub ID'
    ws['B1'] = 'Rater 1'
    ws['C1'] = 'Rater 2'
    for col in range(4, 4 + num_questions):
        ws.cell(column=col, row=1, value="Q{0}".format(col-3))
        question_metadata = ws.cell(column=col,row=2)
        question_metadata.value = '=Entry1!{0}'.format(question_metadata.coordinate)

    for row in range(3, 3 + num_participants):
        subject = ws.cell(column=1,row=row)
        subject.value = "=Entry1!A{0}".format(row)
        rater1 = ws.cell(column=2,row=row)
        rater1.value = "=Entry1!C{0}".format(row)
        rater2 = ws.cell(column=3,row=row)
        rater2.value = "=Entry2!C{0}".format(row)
        for col in range(4, 4 + num_questions):
            cell = ws.cell(column=col,row=row)
            cell.value = "=IF(Entry1!{0}=Entry2!{0},Entry2!{0},CONCATENATE(Entry1!{0},\" vs. \",Entry2!{0}))".format(cell.coordinate)

    last_cell = ws.cell(column=3 + num_questions, row=2 + num_participants)

    # Conditional formatting to highlight mismatches
    redFill = PatternFill(start_color='EE6666', end_color='EE6666', fill_type='solid')
    dxf = DifferentialStyle(fill=redFill)
    rule = Rule(type="containsText", operator="containsText", text="vs.", dxf=dxf)
    ws.conditional_formatting.add('D3:{0}'.format(last_cell.coordinate), rule)


def generate_workbook(name, num_questions, num_participants, range_high, range_low=1):
    wb = Workbook()

    wb.remove(wb.active)
    fill_sheet(wb.create_sheet("Entry1"), num_questions, num_participants, range_high, range_low)
    fill_sheet(wb.create_sheet("Entry2"), num_questions, num_participants, range_high, range_low, copyRow2FromSheet1=True)
    compare_sheet(wb.create_sheet("Final_Comparison"), num_questions, num_participants)

    if not ".xlsx" in name:
        name = name + ".xlsx"
    wb.save(name)


def generate_manual():
    generate_workbook(args.name, args.num_questions, args.num_participants, args.range_high)


def generate_automatic():
    num_participants = args.num_participants

    # load file
    workbook = xlrd.open_workbook(args.name)
    sheet = workbook.sheet_by_index(0)

    for row in range(1,sheet.nrows):
        num_questions = sheet.cell_value(row,1)
        if num_questions == '': continue
        num_questions = int(num_questions)
        if num_questions > 0:
            name = sheet.cell_value(row,0)
            consistent_scale = sheet.cell_value(row,2)
            if consistent_scale == 'y':
                low = sheet.cell_value(row,3)
                high = sheet.cell_value(row,4)
            else:
                low = None
                high = None

            generate_workbook(os.path.join(args.output_path, name), num_questions, num_participants, high, low)


if args.subcommand == 'auto':
    generate_automatic()
else:
    generate_manual()
