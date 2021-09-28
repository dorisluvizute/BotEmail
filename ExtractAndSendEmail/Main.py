import sys
sys.path.insert(0, '/EmailBot')

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
files.append(os.environ["DIRETORIO"] + "Attachments\Formulario_Admissao.docx")
files.append(os.environ["DIRETORIO"] + "Attachments\GDP-FOR-001-067_-_DECLARACAO_DE_DEPENDENTES_PARA_FINS_DE_IMPOSTO_DE_RENDA.docx")
files.append(os.environ["DIRETORIO"] + "Attachments\Lista Documentacao_Qintess.pdf")

emails = SheetService.get_all_emails()
for email in emails:
    print("*Etapa de enviar email*")
    recipients = [email, os.environ["CC"]]
    count = 1
    try:
        if not SheetService.is_already_send(email):
            SendEmailService.send_mail(os.environ["USER"], recipients, os.environ["EMAIL_SUBJECT"], Dictionary.body_text(), files)
            SheetService.change_excel_status(email, "ENVIADO")
            count += 1
    except:
        count += 1
        raise ValueError('Erro ao enviar email.')

                





                
            


