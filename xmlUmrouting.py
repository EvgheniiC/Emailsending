import os
import re
from email.message import EmailMessage
import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import shutil

xlsx_source_path = r'C:\Users\djoni\Downloads\bsp'

def send_email():

    sender_address = 'myEmaail@gmail.com'
    sender_pass = 'mypasswort'
    receiver_address = 'adressTo@identity.sixt.com'

    for file in get_attachments():
        message = MIMEMultipart()
        message['From'] = sender_address
        message['To'] = receiver_address
        message['Subject'] = 'These are files that were sent to the wrong address.'
        mail_content = '''Hello, these are lost accounts'''

        file_pdf = f"{xlsx_source_path}\{file}.pdf"
        file_xml = f"{xlsx_source_path}\{file}.xml"

        with open(file_pdf, 'rb') as f:
            message.attach(MIMEText(mail_content, 'plain'))
            payload = MIMEBase('application', 'octate-stream')
            payload.set_payload(f.read())
            encoders.encode_base64(payload)  # encode the attachment
            payload.add_header('Content-Disposition', 'attachment', filename=f"{file}.pdf")
            message.attach(payload)

        with open(file_xml, 'rb') as f:
            message.attach(MIMEText(mail_content, 'plain'))
            payload = MIMEBase('application', 'octate-stream')
            payload.set_payload(f.read())
            encoders.encode_base64(payload) 
            payload.add_header('Content-Disposition', 'attachment', filename=f"{file}.xml")
            message.attach(payload)
        
        # Create SMTP session for sending the mail
        try:
            session = smtplib.SMTP('smtp.gmail.com', 587)  # use gmail with port
            session.starttls()  # enable security
            session.login(sender_address, sender_pass)  # login with mail_id and password
            text = message.as_string()
            session.sendmail(sender_address, receiver_address, text)
            session.quit()

            processed = r'C:\Users\djoni\Downloads\processed'
            shutil.move(file_pdf,f"{processed}\{file}.pdf")
            shutil.move(file_xml,f"{processed}\{file}.xml")
            print('Mail Sent succesfull')
        except Exception as e:
            error = r'C:\Users\djoni\Downloads\error'
            shutil.move(file_pdf,f"{error}\{file}.pdf")
            shutil.move(file_xml,f"{error}\{file}.xml")
            print('Mail Not sended')
            print(e)


def get_attachments():
    files = [re.sub(".xml|.pdf", "", file)
             for file in os.listdir(xlsx_source_path)]
    return list(dict.fromkeys(files))

send_email()
