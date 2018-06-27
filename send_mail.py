import smtplib, os
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.utils import COMMASPACE, formatdate
from email import encoders


class MailServerConfig:

    def __init__(self, host, port, tls, html):
        self.host = host
        self.port = port
        self.tls = tls
        self.html = html


class AvailableServers:

    MICROSOFT_EXCHANGE = MailServerConfig("smtp.office365.com", port=587, tls=True, html=False)


class Credentials:

    def __init__(self, username, password):
        self.username = username
        self.password = password


def send_mail(sender, receivers, subject, body, files=[], credentials=None, server=AvailableServers.MICROSOFT_EXCHANGE):

    msg = MIMEMultipart('related')
    msg['From'] = sender
    msg['To'] = COMMASPACE.join(receivers)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach( MIMEText(body, 'html' if server.html else 'plain') )

    for f in files:
        part = MIMEBase('application', "octet-stream")
        part.set_payload( open(f,"rb").read() )
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(f))
        msg.attach(part)

    print(server.host, server.port)
    smtp = smtplib.SMTP(server.host, server.port)
    if server.tls:
        smtp.starttls()

    if credentials is not None:
        smtp.login(credentials.username, credentials.password)

    smtp.sendmail(sender, receivers, msg.as_string())
    smtp.close()

