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
manual.add_argument('num_participants', type=int, help='How many participant rows to generate')
manual.add_argument('--range-high', dest='range_high', type=int, help='If there is a range for all questions, what is the top of that range for conditional formatting?')

auto = subparsers.add_parser('auto')
auto.add_argument('input', help='Input excel file that specifies what questionairres to create')
auto.add_argument('num_participants', type=int, help='How many participant rows to generate')
auto.add_argument('output_path', help='Where to create the files')

args = parser.parse_args()

small_font = Font(size = "8")
border_right = Border(right=Side(style='double', color='0000ff'))
border_bottom = Border(bottom=Side(style='double', color='0000ff'))


def do_borders(ws, cols, rows):
    for col in range(6, cols):
        ws.cell(column=col, row=3).border = border_bottom
    for row in range(4, rows):
        ws.cell(column=5, row=row).border = border_right

def fill_sheet(ws, num_questions, num_participants, range_high, range_low=1, copyRowsFromSheet1=False):
    # Add headers
    ws['A1'] = 'Sub ID'
    ws['B1'] = 'Session Date'
    ws['C1'] = 'Visit'
    ws['D1'] = 'Date Entered'
    ws['E1'] = 'Entered By'
    for col in range(1, 6 + num_questions):
        ws.column_dimensions[get_column_letter(col)].width = "15"
    for col in range(6, 6 + num_questions):
        question_metadata1 = ws.cell(column=col,row=2)
        question_metadata1.font = small_font
        question_metadata2 = ws.cell(column=col,row=3)
        question_metadata2.font = small_font
        if copyRowsFromSheet1:
            ws.cell(column=col, row=1, value='=Entry1!{0}{1}'.format(get_column_letter(col), 1))
            question_metadata1.value='=Entry1!{0}{1}'.format(get_column_letter(col), 2)
            question_metadata2.value='=Entry1!{0}{1}'.format(get_column_letter(col), 3)
        else:
            ws.cell(column=col, row=1, value="Q{0}".format(col-5))

    ws.cell(column=6 + num_questions, row=1, value="Notes")

    # Add subject numbers
    for row in range(4, 4 + num_participants):
        ws.cell(column=1, row=row, value=(row-3+1000))

    # This freezes everything "above and to the left" of the given cell
    ws.freeze_panes = "F4"

    do_borders(ws, 6 + num_questions, 4 + num_participants)
    
    last_cell = ws.cell(column=5 + num_questions, row=2 + num_participants)

    # Conditional formatting to do blanks as light gray
    blankFill = PatternFill(start_color='EEEEEE', end_color='EEEEEE', fill_type='solid')
    dxf = DifferentialStyle(fill=blankFill)
    blank = Rule(type="expression", formula = ["ISBLANK(F4)"], dxf=dxf, stopIfTrue=True)
    ws.conditional_formatting.add('F4:{0}'.format(last_cell.coordinate), blank)
                              
    # Conditional formatting to highlight numbers out of range
    if range_high:
        yellowFill = PatternFill(start_color='EEEE66', end_color='EEEE66', fill_type='solid')
        rule = CellIsRule(operator="notBetween", formula=[str(range_low),str(range_high)], fill=yellowFill)
        ws.conditional_formatting.add('F4:{0}'.format(last_cell.coordinate), rule)

    

def compare_sheet(ws, num_questions, num_participants):
    # Protect the sheet so nobody can edit it
    ws.protection.sheet = True
    # Add headers
    ws['A1'] = 'Sub ID'
    ws['B1'] = 'Session Date'
    ws['C1'] = 'Visit'
    ws['D1'] = 'Rater 1'
    ws['E1'] = 'Rater 2'
    for col in range(1, 6 + num_questions):
        ws.column_dimensions[get_column_letter(col)].width = "15"
    for col in range(6, 6 + num_questions):
        ws.cell(column=col, row=1, value="Q{0}".format(col-5))
        question_metadata1 = ws.cell(column=col,row=2)
        question_metadata1.value = '=Entry1!{0}{1}'.format(get_column_letter(col), 2)
        question_metadata1.font = small_font
        question_metadata2 = ws.cell(column=col,row=3)
        question_metadata2.value = '=Entry1!{0}{1}'.format(get_column_letter(col), 3)
        question_metadata2.font = small_font

    def compare_cell(cell):
        cell.value = "=IF(Entry1!{0}=Entry2!{0},Entry2!{0},CONCATENATE(Entry1!{0},\" vs. \",Entry2!{0}))".format(cell.coordinate)

    for row in range(4, 4 + num_participants):
        subject = ws.cell(column=1,row=row)
        compare_cell(subject)
        session_date = ws.cell(column=2,row=row)
        compare_cell(session_date)
        visit = ws.cell(column=3,row=row)
        compare_cell(visit)
        rater1 = ws.cell(column=4,row=row)
        rater1.value = "=Entry1!E{0}".format(row)
        rater2 = ws.cell(column=5,row=row)
        rater2.value = "=Entry2!E{0}".format(row)
        for col in range(4, 6 + num_questions):
            cell = ws.cell(column=col,row=row)
            compare_cell(cell)

    last_cell = ws.cell(column=3 + num_questions, row=3 + num_participants)

    for col in range(4, 6 + num_questions):
        ws.column_dimensions[get_column_letter(col)].width = "15"

    # Conditional formatting to highlight mismatches
    redFill = PatternFill(start_color='EE6666', end_color='EE6666', fill_type='solid')
    dxf = DifferentialStyle(fill=redFill)
    rule = Rule(type="containsText", operator="containsText", text=" vs. ", dxf=dxf)
    ws.conditional_formatting.add('A3:{0}'.format(last_cell.coordinate), rule)


def compare_sheet_vertical(ws, num_questions, num_participants):
    # Add headers
    ws['A1'] = 'Sub ID'
    ws['B1'] = 'Question'
    ws['C1'] = 'Rater'
    ws['D1'] = 'Value'


def generate_workbook(name, num_questions, num_participants, range_high, range_low=1):
    wb = Workbook()

    wb.remove(wb.active)
    fill_sheet(wb.create_sheet("Entry1"), num_questions, num_participants, range_high, range_low)
    fill_sheet(wb.create_sheet("Entry2"), num_questions, num_participants, range_high, range_low, copyRowsFromSheet1=True)
    compare_sheet(wb.create_sheet("Final_Comparison"), num_questions, num_participants)
    #compare_sheet_vertical(wb.create_sheet("Vertical_Comparison"), num_questions, num_participants)

    if not ".xlsx" in name:
        name = name + ".xlsx"
    wb.save(name)


def generate_manual():
    generate_workbook(args.name, args.num_questions, args.num_participants, args.range_high)


def generate_automatic():
    num_participants = args.num_participants

    # load file
    workbook = xlrd.open_workbook(args.input)
    sheet = workbook.sheet_by_index(0)

    for row in range(1,sheet.nrows):
        num_questions = sheet.cell_value(row,4)
        if num_questions == '': continue
        try:
            num_questions = int(num_questions)
        except ValueError:
            print(f'Ignoring row {row}, "{num_questions}" is not a question number I understand')
            continue
        if num_questions > 0:
            name = sheet.cell_value(row,1)
            consistent_scale = sheet.cell_value(row,5)
            if consistent_scale == 'y':
                low = sheet.cell_value(row,6)
                high = sheet.cell_value(row,7)
            else:
                low = None
                high = None

            if low == "":
                low = None
            if high == "":
                high = None

            generate_workbook(os.path.join(args.output_path, name), num_questions, num_participants, high, low)


if args.subcommand == 'auto':
    generate_automatic()
else:
    generate_manual()
