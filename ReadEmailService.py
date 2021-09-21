import imaplib
import email

from email.header import decode_header
from typing import cast
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def read_email(subject_filter):
    username = "dluvizute@hotmail.com"
    password = "Doris1509*"

    # create an IMAP4 class with SSL 
    imap = imaplib.IMAP4_SSL("outlook.office365.com")

    # authenticate
    imap.login(username, password)

    status, messages = imap.select("INBOX")
    # number of top emails to fetch
    N = 200
    # total number of emails
    messages = int(messages[0])
    print("Total de emails - %s" % messages)

    print(status)

    mail_content = []
    for i in range(messages, messages-N, -1):
        # fetch the email message by ID
        res, msg = imap.fetch(str(i), "(RFC822)")
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
                    
                    #subject = subject.decode()

                # email sender
                from_ = msg.get("From")
                # print("Subject:", subject)
                # print("From:", from_)

                if subject_filter in subject:
                    if 'Não é possível entregar' in subject_filter:
                        print("--- EMAIL DE ERRO LOCALIZADO ---")
                        if msg.is_multipart():
                            mail_content.append(msg.get_payload()[0].get_payload()[0].get_payload())

                    elif 'Informações Cadastrais para Depto Pessoal' in subject_filter:
                        print("--- FO LOCALIZADA ---")                    
                        if msg.is_multipart():
                            for part in msg.get_payload():
                                if part.get_content_type() == 'text/plain':
                                    mail_content.append(part.get_payload())                

    
    return mail_content

