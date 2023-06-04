from config import sender_email, app_password, email_subject

import filetype
from smtplib import SMTP_SSL
from email.message import EmailMessage

# Gmail Server configration
smtp_server_url: str = "smtp.gmail.com"
smtp_server_port: int = 465


def send_email_messages(email_messages: list[EmailMessage]):
    with SMTP_SSL(smtp_server_url, smtp_server_port) as smtp:
        smtp.login(sender_email, app_password)
        for email_message in email_messages:
            smtp.send_message(email_message)


def get_email_message(receiver_address: str, body: str, attachment_path: str):
    # Email Message object which is used to send messages
    email_message = EmailMessage()
    email_message['From'] = sender_email
    email_message['To'] = receiver_address
    email_message['Subject'] = email_subject

    # Reading binary data from the attachment to add in the email_message
    with open(attachment_path, 'rb') as attachment:
        attachment_data = attachment.read()
        attachment_type = filetype.guess(attachment.name)

    # Setting email message body and adding the photo attachment
    email_message.set_content(body)
    email_message.add_attachment(attachment_data, maintype=attachment_type.mime, filename=attachment.name,
                                 subtype=attachment_type.extension)
    return email_message
