import smtplib, os, traceback
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.utils import COMMASPACE, formatdate
from email import encoders
import config


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

try:
    import win32com.client as win32
    USING_OUTLOOK_STATUS = True

    def send_mail_by_outlook(sender, receipients, subject, message, attachments):        
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = ";".join(receipients)
        mail.Subject = subject
        mail.Body = message
        for path in attachments:
            mail.Attachments.Add(path)
        mail.Send()

except:
    print("cannot access outlook api")
    print("will try to use direct mail server access")
    USING_OUTLOOK_STATUS = False

def send_mail_direct(sender, receivers, subject, body, files=[], credentials=None, server=AvailableServers.MICROSOFT_EXCHANGE):

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

    smtp = smtplib.SMTP(server.host, server.port)
    if server.tls:
        smtp.starttls()

    if credentials is not None:
        smtp.login(credentials.username, credentials.password)

    smtp.sendmail(sender, receivers, msg.as_string())
    smtp.close()


def send_mail(sender, receivers, subject, message, attachments, **kwargs):
    print("--------------")
    print("sending email:")
    print("from: {}".format(sender))
    print("to: {}".format(str(receivers)))
    print("subject: {}".format(subject))
    print("message: {}".format(message))
    print("attach: {}".format(str(attachments)))
    print("--------------")

    if USING_OUTLOOK_STATUS and config.use_outlook:
        send_mail_by_outlook(sender, receivers, subject, message, attachments)

    else:
        try:
            credentials = kwargs["credentials"]
            #server = kwargs["server"]
            send_mail_direct(sender, receivers, subject, message, attachments, credentials)

        except Exception as e:
            print("sending mail direct failed...")
            traceback.print_exc(e)

