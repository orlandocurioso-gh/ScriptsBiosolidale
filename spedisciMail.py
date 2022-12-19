import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from email.mime.application import MIMEApplication


def mail(to, subject,text,nomefile):
    gmail_user = "felice@biosolidale.it"
    gmail_pwd = "BioFelice"
    msg = MIMEMultipart()
    msg['From'] = gmail_user
    msg['To'] = to
    msg['Subject'] = subject
    msg.attach(MIMEText(text))
    mail_file = MIMEBase('application', 'csv')
    mail_file.set_payload(open('C://Users//f.altarocca//Documents//coding//Python//scripts38//ListiniAManoSender//Listini//'+nomefile, 'rb').read())
    mail_file.add_header('Content-Disposition', 'attachment', filename=nomefile)
    encoders.encode_base64(mail_file)
    msg.attach(mail_file)
    mailServer = smtplib.SMTP("smtp.gmail.com", 587)
    mailServer.ehlo()
    mailServer.starttls()
    mailServer.ehlo()
    mailServer.login(gmail_user, gmail_pwd)
    risposta=mailServer.sendmail(gmail_user, to, msg.as_string())
    mailServer.close()
    return risposta