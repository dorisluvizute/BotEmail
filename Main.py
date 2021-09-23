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

    email_employee = employee_data[14].split("<")[0].strip()
    
    if not SheetService.check_if_email_exists_in_worksheet(email_employee):
        name_employee = employee_data[12].strip()         
        name_employee = Dictionary.convert_ISO8859(name_employee, dict)

        SheetService.fill_excel_worksheet(email_employee, name_employee)


files  =  []
files.append(os.environ["DIRETORIO"] + "Formulario_Admissao.docx")
files.append(os.environ["DIRETORIO"] + "GDP-FOR-001-067_-_DECLARACAO_DE_DEPENDENTES_PARA_FINS_DE_IMPOSTO_DE_RENDA.docx")
files.append(os.environ["DIRETORIO"] + "Lista Documentacao_Qintess.pdf")

# files.append("C:\\EmailBot\\files\\Formulario_Admissao.docx")
# files.append("C:\\EmailBot\\files\\GDP-FOR-001-067_-_DECLARACAO_DE_DEPENDENTES_PARA_FINS_DE_IMPOSTO_DE_RENDA.docx")
# files.append("C:\\EmailBot\\files\\Lista Documentacao_Qintess.pdf")

emails = SheetService.get_all_emails()
for email in emails:
    count = 1
    try:
        if not SheetService.is_already_send(email):
            SendEmailService.send_mail(os.environ["USER"], email, os.environ["EMAIL_SUBJECT"], Dictionary.body_text(), files)
            # SendEmailService.send_mail("dluvizute@hotmail.com", email, "BEM VINDO - PROCESSO DE ADMISSAO QINTESS", Dictionary.body_text(), files)
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

                





                
            


