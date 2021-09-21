from os import error
import ReadEmailService
import SheetService
import SendEmailService

mail_content = ReadEmailService.read_email('Informações Cadastrais para Depto Pessoal')

for email in mail_content:
    employee_data = email.split("\n")

    email_employee = employee_data[14].split("<")[0]
    
    if not SheetService.check_if_email_exists_in_worksheet(email_employee):
        name_employee = employee_data[12].strip()
        SheetService.fill_excel_worksheet(email_employee, name_employee)

text = "Teste enviado pelo bot"

files  =  []
files.append("C:\\EmailBot\\files\\Carta Proposta (Offer Letter) Qintess - Dóris Andressa Moura Luvizute.pdf")

emails = SheetService.get_all_emails()
for email in emails:
    count = 1
    try:
        if not SheetService.is_already_send(email):
            SendEmailService.send_mail("dluvizute@hotmail.com", email, "Teste", text, files)
            SheetService.change_excel_status("ENVIADO")
            count += 1
    except:
        count += 1
        raise ValueError('Erro ao enviar email.')

error_mail_content = ReadEmailService.read_email('Não é possível entregar')
if error_mail_content != []:
    for error in error_mail_content:
        error_mail = error.split(":")[1].split("(")[0].strip()
        count = 1
        for email in emails:            
            if email == error_mail:
                SheetService.change_excel_status("Email inválido")
                count += 1
            else:
                count += 1

                





                
            


