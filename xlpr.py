import argparse
from openpyxl import Workbook

parser = argparse.ArgumentParser(description='Automate creation of scoring Excel worksheets.')
parser.add_argument('name', help='What to call the file')
parser.add_argument('num_questions', type=int, help='How many question columns to generate')
parser.add_argument('num_participants', type=int, help='How many question columns to generate')
args = parser.parse_args()

wb = Workbook()
num_questions = args.num_questions

def fill_sheet(ws):
    ws['A1'] = 'Sub ID'
    ws['B1'] = 'Date'
    ws['C1'] = 'Entered By'

    for col in range(4, 4 + num_questions):
        ws.cell(column=col, row=1, value="Q{0}".format(col-3))

    ws.cell(5 + num_questions, row=1, value="Notes")
    # TODO: Freeze row 1, columns 1-3, add borders
    # Maybe notes should be 4 and maybe frozen?

def compare_sheet(ws):
    cell = ws.cell(column=i,row=j)
    # TODO: how to find cell name for formula
    cell.set_explicit_value(value="=IF($Entry1.C2=$Entry2.C2,$Entry2.C2,'ENTRIES DON’T MATCH')",data_type=cell.TYPE_FORMULA)

wb.remove(wb.active)
fill_sheet(wb.create_sheet("Entry1"))
fill_sheet(wb.create_sheet("Entry2"))
compare_sheet(wb.create_sheet("Final_Comparison"))

# Save the file
wb.save(args.name + ".xlsx")

