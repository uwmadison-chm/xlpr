from openpyxl import Workbook


wb = Workbook()
num_questions = 20

def fill_sheet(ws):
    ws['A1'] = 'PID'
    ws['B1'] = 'Rater initials'
    ws['C1'] = 'Other stuff'

    for col in range(4, 4 + num_questions):
        ws.cell(column=col, row=1, value="Q{0}".format(col-3))

fill_sheet(wb.create_sheet("Rater 1"))
fill_sheet(wb.create_sheet("Rater 2"))

# Save the file
wb.save("output.xlsx")

