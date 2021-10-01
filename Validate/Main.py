import sys
# sys.path.insert(0, 'Admissoes/EmailBot')
sys.path.insert(0, '/EmailBot')

import ReadEmailService
import SheetService

erros = 0

emails = SheetService.get_all_emails()
error_mail_content = ReadEmailService.read_email("Não é possível entregar")
if error_mail_content != []:
    for error in error_mail_content:
        for email in emails:            
            if email == error:
                SheetService.change_excel_status(email, "Email inválido")
                erros += 1

print(erros)