from os import error
import ReadEmailService
import SheetService
import SendEmailService

mail_content = ReadEmailService.read_email()

for email in mail_content:
    count = 1

    employee_data = email.split(":")

    email_employee = employee_data[1].replace("nome", "").strip()
    name_employee = employee_data[2].strip() 

    SheetService.fill_excel_worksheet(count, email_employee, name_employee)
    count += 1

text = "Teste enviado pelo bot"

files  =  []
files.append("C:\\EmailBot\\files\\Carta Proposta (Offer Letter) Qintess - Dóris Andressa Moura Luvizute.pdf")

emails = SheetService.get_all_emails()
for email in emails:
    count = 1
    try:
        SendEmailService.send_mail("dluvizute@hotmail.com", email, "Teste", text, files)
        SheetService.change_excel_status(count)
    except:
        raise ValueError('Erro ao enviar email.')

error_mail_content = ReadEmailService.find_error_email()
if error_mail_content != []:
    print(error_mail_content)





                
            


