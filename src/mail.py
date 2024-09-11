import email.utils
import logging
import os
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import List

from src.data import Mail


def send_mail(mail_info: Mail) -> bool:
    recipients_lst: List[str] = mail_info.recipients.split(";")

    msg = MIMEMultipart()
    msg["From"] = mail_info.sender
    msg["To"] = mail_info.recipients
    msg["Date"] = email.utils.formatdate(localtime=True)
    msg["Subject"] = mail_info.subject
    msg.attach(MIMEText(mail_info.subject, "html", "utf-8"))

    attachment_name = os.path.basename(mail_info.attachment_path)
    with open(mail_info.attachment_path, "rb") as f:
        part = MIMEApplication(f.read(), Name=attachment_name)
    part.add_header("Content-Disposition", "attachment", filename=attachment_name)
    msg.attach(part)

    try:
        with smtplib.SMTP(mail_info.server, 25) as smtp:
            response = smtp.sendmail(mail_info.sender, recipients_lst, msg.as_string())
            if response:
                logging.error("Failed to send email to the following recipients:")
                for recipient, error in response.items():
                    logging.error(f"{recipient}: {error}")
                return False
            else:
                logging.info("Email sent successfully.")
                return True
    except smtplib.SMTPException as e:
        logging.error(f"Failed to send email: {e}")
        return False
