import gspread
import os, sys
from docxtpl import DocxTemplate


os.chdir(sys.path[0])

spreadsheet_name = "complaints_PFU"
worksheet_name = "AZ_execute_decisions"
docx_template_name = "../Template_az.docx"


def render_doc(row_data, template_name, work_sheet, row_number):
    doc = DocxTemplate(template_name)
    context = row_data
    doc.render(context)
    filename = "../"+row.get("number")+" "+row.get("doc_name")+".docx"
    doc.save(filename)
    work_sheet.update_cell(row_number, 1, 1)

sa = gspread.service_account(filename = "../service_account.json")
sh = sa.open(spreadsheet_name)
wks = sh.worksheet(worksheet_name)

number_of_row = 2
for row in wks.get_all_records():
    print(number_of_row)
    if row.get("done") == 0:
        render_doc(row, docx_template_name, wks, number_of_row)
        print(f"Сгенеровано файл: {row.get('number')}{row.get('doc_name')}.docx")
    number_of_row = number_of_row + 1
