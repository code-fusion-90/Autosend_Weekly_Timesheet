import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import date
import os.path

## Sends email along with attachment at path given as attachment_location. For example attachment_location = /Users/Timesheet.xls
## Timesheet.xls will be attached and sent over to the recipients.
## Replace youremail, r1, r2, r3 with reciepients email-id & replace mailpwd with email password.
## Errors not handled.
def email(attachment_location):
    today = date.today()
    email_sender = 'r1'
    msg = MIMEMultipart()
    msg['From'] = 'youremail'
    to_emails = ["r1", "r2"]
    cc_emails = ["r3", "r4"]
    msg["To"] = ', '.join(to_emails)
    msg["cc"] = ', '.join(cc_emails)
    #bcc_emails = [""]
    emails = to_emails + cc_emails
    msg['Subject'] = 'Weekly Time Sheet - ' + str(today)
    body = """
    Hi All,

    PFA the time sheet for the week ending {str1}

    Thanks & Regards
    --Your Name
    """.format(str1 = str(today))
    msg.attach(MIMEText(body))
    
    if attachment_location != '':
        filename = os.path.basename(attachment_location)
        attachment = open(attachment_location, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
        msg.attach(part)

    mailserver = smtplib.SMTP('smtp.office365.com',587)
    mailserver.ehlo()
    mailserver.starttls()
    mailserver.login('yourmail', 'mailpwd')
    text = msg.as_string()
    mailserver.sendmail('yourmail', emails, text)
    mailserver.quit()
    print('Mail Sent Successfully')