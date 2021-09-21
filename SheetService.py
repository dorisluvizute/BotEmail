import email
from openpyxl import load_workbook

def fill_excel_worksheet(email_employee, name_employee):
    try:
        count = 1
        wb = load_workbook("C:\\EmailBot\\files\\PlanilhaNovosFuncionarios.xlsx")
        ws = wb.worksheets[0]

        while ws["A" + str(count+1)].value != None and ws["B" + str(count+1)].value != None:
            count += 1

        ws["A" + str(count+1)] = name_employee
        ws["B" + str(count+1)] = email_employee

        wb.save("C:\\EmailBot\\files\\PlanilhaNovosFuncionarios.xlsx")
    except:
        raise ValueError('Não foi possível preencher a planilha com os dados do empregado.')

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

def change_excel_status(status):
    try:
        count = 1
        wb = load_workbook("C:\\EmailBot\\files\\PlanilhaNovosFuncionarios.xlsx")
        ws = wb.worksheets[0]

        ws["C" + str(count+1)] = status

        wb.save("C:\\EmailBot\\files\\PlanilhaNovosFuncionarios.xlsx")
    
    except:
        raise ValueError("Não foi possível alterar o status no excel.")


def check_if_email_exists_in_worksheet(email):
    wb = load_workbook("C:\\EmailBot\\files\\PlanilhaNovosFuncionarios.xlsx")
    ws = wb.worksheets[0]

    validate = False
    for cell in ws["B"]:
        if cell.value == email:
            validate = True 
    
    return validate

def is_already_send(email):
    wb = load_workbook("C:\\EmailBot\\files\\PlanilhaNovosFuncionarios.xlsx")
    ws = wb.worksheets[0]

    count = 1
    status = ''
    for cell in ws["B"]:
        if cell.value == email:
            status = ws["C" + str(count+1)].value

        count += 1
    
    if status == "ENVIADO":
        return True

    return False
