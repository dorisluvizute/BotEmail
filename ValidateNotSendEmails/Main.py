import sys
sys.path.insert(0, '/EmailBot')

import ReadEmailService
import SheetService

emails = SheetService.get_all_emails()
error_mail_content = ReadEmailService.read_email("Não é possível entregar")
if error_mail_content != []:
    for error in error_mail_content:
        count = 1
        for email in emails:            
            if email == error:
                SheetService.change_excel_status(email, "Email inválido")
                count += 1
            else:
                count += 1