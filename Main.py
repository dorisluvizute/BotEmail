from os import error
import os
import ReadEmailService
import SheetService
import SendEmailService
import Dictionary

dict = Dictionary.dictionary()
mail_content = ReadEmailService.read_email("Informações Cadastrais para Depto Pessoal")

for email in mail_content:
    email_employee = email[0].get_payload()[1].get_payload().split("mailto:")[1].split('"')[0]
    
    if not SheetService.check_if_email_exists_in_worksheet(email_employee):
        name_employee = email[0].get_payload()[1].get_payload().split("Nome")[1].split(">")[7].replace("</b", "")         

        if "&nbsp;" in name_employee:
            name_employee = name_employee.replace("&nbsp;", "") 
            
        name_employee = Dictionary.convert_ISO8859(name_employee, dict)

        SheetService.fill_excel_worksheet(email_employee, name_employee)


files  =  []
files.append(os.environ["DIRETORIO"] + "Formulario_Admissao.docx")
files.append(os.environ["DIRETORIO"] + "GDP-FOR-001-067_-_DECLARACAO_DE_DEPENDENTES_PARA_FINS_DE_IMPOSTO_DE_RENDA.docx")
files.append(os.environ["DIRETORIO"] + "Lista Documentacao_Qintess.pdf")

emails = SheetService.get_all_emails()
for email in emails:
    count = 1
    try:
        if not SheetService.is_already_send(email):
            SendEmailService.send_mail(os.environ["USER"], email, os.environ["EMAIL_SUBJECT"], Dictionary.body_text(), files)
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

                





                
            


