import win32com.client as win32
import config

def send_mail(sender, receipients, subject, message, attachments):
    print("--------------")
    print("sending email:")
    print("from: {}".format(sender))
    print("to: {}".format(str(receipients)))
    print("subject: {}".format(subject))
    print("message: {}".format(message))
    print("attach: {}".format(str(attachments)))
    print("--------------")
    
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = ";".join(receipients)
    mail.Subject = subject
    mail.Body = message
    #mail.HTMLBody = '<h2>HTML Message body</h2>'
    for path in attachments:
        mail.Attachments.Add(path)
    mail.Send()
