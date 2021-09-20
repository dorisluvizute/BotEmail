import email
from openpyxl import load_workbook

def fill_excel_worksheet(count, email_employee, name_employee):
    wb = load_workbook("C:\\EmailBot\\files\\PlanilhaNovosFuncionarios.xlsx")
    ws = wb.worksheets[0]

    ws["A" + str(count+1)] = name_employee
    ws["B" + str(count+1)] = email_employee
    wb.save("C:\\EmailBot\\files\\PlanilhaNovosFuncionarios.xlsx")

def get_all_emails():
    wb = load_workbook("C:\\EmailBot\\files\PlanilhaNovosFuncionarios.xlsx")
    ws = wb.worksheets[0]

    emails = []
    for cell in ws['B']:
        if cell.value == 'Email' or cell.value == None:
            continue
        else:
            emails.append(cell.value)
    
    return emails

def change_excel_status(count):
    wb = load_workbook("C:\\EmailBot\\files\\PlanilhaNovosFuncionarios.xlsx")
    ws = wb.worksheets[0]

    ws["C" + str(count+1)] = "ENVIADO"
    wb.save("C:\\EmailBot\\files\\PlanilhaNovosFuncionarios.xlsx")