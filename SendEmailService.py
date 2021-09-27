import smtplib, ssl
import os
import DecryptService

from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate


def send_mail(send_from, send_to, subject, text, files=[]):
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = send_to
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    text = text.split("\n")
    html = '<html><head><style> #orientacao {font-size: 16px;} .chamativo {font-size: 16px; font-style: bold; color: #9100F5} .importante {font-style: bold; color: #9100F5}</style></head><body>'
    html += f'<br>{text[0]}<br><br>{text[1]}<br><br>{text[2]}<br><br>{text[3]}<br>'
    html += f'<br><u id = "orientacao">{text[4]}</u>'
    html += f'<br><strong class = "chamativo">{text[5]}</strong>'
    html += f'<br>{text[6]}<br>{text[7]}<br>{text[8]}<br>'
    html += f'<br><strong class = "chamativo">{text[9]}</strong>'
    html += f'<br>{text[10]}<br>'
    html += f'<br><strong class = "importante"><center>{text[11]}</center></strong>'
    html += f'<br><strong class = "importante">{text[12]}</strong>'
    html += f'</body></html>'
  

    msg.attach(MIMEText(html, 'html'))

    for f in files or []:
        with open(f, "rb") as fil:
            part = MIMEApplication(
                fil.read(),
                Name=basename(f)
            )
        # After the file is closed
        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
        msg.attach(part)

    context = ssl.create_default_context()

    smtp = smtplib.SMTP("smtp.office365.com", 587)
    smtp.ehlo()
    
    smtp.starttls(context=context)
    smtp.login(os.environ["USER"], DecryptService.decrypt(os.environ["PASSWORD"]))
    
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.close()

