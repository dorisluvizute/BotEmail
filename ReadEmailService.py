import imaplib
import email
import os
import DecryptService

from email.header import decode_header
from typing import cast
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def read_email(subject_filter):
    # username = os.environ["USER"]
    # password = DecryptService.decrypt(os.environ["PASSWORD"])
    username = "bot@resource.com.br"
    password = DecryptService.decrypt("UWludGVzc0AyMDIx")

    # create an IMAP4 class with SSL 
    imap = imaplib.IMAP4_SSL("outlook.office365.com")
    
    # authenticate
    imap.login(username, password)
    status, messages = imap.select("INBOX")

    # number of top emails to fetch
    N = 200
    
    # total number of emails
    messages = int(messages[0])

    mail_content = []
    for i in range(messages, messages-N, -1):
        # fetch the email message by ID
        try:
            res, msg = imap.fetch(str(i), "(RFC822)")
        except:
            break
        
        for response in msg:
            if isinstance(response, tuple):
                # parse a bytes email into a message object
                msg = email.message_from_bytes(response[1])

                # decode the email subject
                subject = decode_header(msg["Subject"])[0][0]
                if isinstance(subject, bytes):

                    # if it's a bytes, decode to str
                    #decode para o email do bpms
                    try:
                        subject = subject.decode(encoding='UTF-8')
                    except:                    
                        subject = subject.decode(encoding='ISO-8859-1')

                # email sender
                from_ = msg.get("From")

                if subject_filter in subject or "Undeliverable" in subject:
                    if 'Não é possível entregar' in subject_filter:
                        if msg.is_multipart():
                            if "&quot;" in msg.get_payload()[0].get_payload().split("To: ")[1].split(",")[0]:
                                mail_content.append(msg.get_payload()[0].get_payload().split("To: ")[1].split(",")[0].split("&quot;")[1])
                            else:
                                msg.get_payload()[0].get_payload().split("To: ")[1].split(",")[0]

                    elif 'Informações Cadastrais para Depto Pessoal' in subject_filter:                
                        if msg.get_default_type() == 'text/plain':
                            mail_content.append(msg.get_payload())              

    
    return mail_content

def delete_all_emails(subject_filter):
    # username = os.environ["USER"]
    # password = DecryptService.decrypt(os.environ["PASSWORD"])
    username = "bot@resource.com.br"
    password = DecryptService.decrypt("UWludGVzc0AyMDIx")

    # create an IMAP4 class with SSL 
    imap = imaplib.IMAP4_SSL("outlook.office365.com")
    
    # authenticate
    imap.login(username, password)
    status, messages = imap.select("INBOX")

    status, messages = imap.search(None, 'FROM "' + subject_filter + '"')

    messages = messages[0].split(b' ')
    for mail in messages:
        _, msg = imap.fetch(mail, "(RFC822)")

        for response in msg:
            if isinstance(response, tuple):
                msg = email.message_from_bytes(response[1])
                subject = decode_header(msg["Subject"])[0][0]

                if isinstance(subject, bytes):
                    try:
                        subject = subject.decode(encoding='UTF-8')
                    except:                    
                        subject = subject.decode(encoding='ISO-8859-1')

        imap.store(mail, "+FLAGS", "\\Deleted")

    imap.expunge()
    imap.close()
    imap.logout()

