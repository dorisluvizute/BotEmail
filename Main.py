from os import error
import os
import ReadEmailService
import SheetService
import SendEmailService
import Dictionary

dict = Dictionary.dictionary()
mail_content = ReadEmailService.read_email("Informações Cadastrais para Depto Pessoal")

for email in mail_content:
    employee_data = email.split("\n")

    email_employee = employee_data[14].split("<")[0]
    
    if not SheetService.check_if_email_exists_in_worksheet(email_employee):
        name_employee = employee_data[12].strip()         
        name_employee = Dictionary.convertISO8859(name_employee, dict)

        SheetService.fill_excel_worksheet(email_employee, name_employee)

text = "Olá seja bem-Vindo!!!\n" + "Parabéns! Você agora faz parte da Qintess, empresa com trajetória de sucesso e grandes realizações na área de Tecnologia da Informação que nos posiciona hoje como um dos principais players em Transformação de Negócios da América Latina.\n"


files  =  []
files.append(os.environ["FILE_ONE"])
files.append(os.environ["FILE_TWO"])
files.append(os.environ["FILE_THREE"])

emails = SheetService.get_all_emails()
for email in emails:
    count = 1
    try:
        if not SheetService.is_already_send(email):
            SendEmailService.send_mail(os.environ["USER"], email, "Teste", text, files)
            SheetService.change_excel_status(email, "ENVIADO")
            count += 1
    except:
        count += 1
        raise ValueError('Erro ao enviar email.')

error_mail_content = ReadEmailService.read_email("Não é possível entregar")
if error_mail_content != []:
    for error in error_mail_content:
        error_mail = error.split(":")[1].split("(")[0].split("<")[0].strip()
        count = 1
        for email in emails:            
            if email == error_mail:
                SheetService.change_excel_status(email, "Email inválido")
                count += 1
            else:
                count += 1

                





                
            


