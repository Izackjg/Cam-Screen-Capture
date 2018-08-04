import ntpath
import os
import smtplib
import winshell
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

class EmailSender:
    def __init__(self, host, port, address, password):
        self.host = host
        self.port = port
        self.address = address
        self.password = password

    def send_message(self, message="This is a message", email_to=None, email_from=None, dt_file_path=None, attach=False):
        if email_to is None:
            email_to = self.address
        if email_from is None:
            email_from = self.address
        if dt_file_path is None:
            captures_folder = winshell.desktop(common=False) + "\\Captures"
            dt = datetime.now().strftime("%B %Y")
            dt_file_path = captures_folder + "\\" + dt

        smtp = smtplib.SMTP(host=self.host, port=self.port)
        smtp.starttls()
        smtp.login(self.address, self.password)
        name = self.split_name_from_address(email_to)
        msg = MIMEMultipart()
        msg["From"] = email_from
        msg["To"] = email_to
        msg["Subject"] = "Videos"

        msg.attach(MIMEText(message, "plain"))
        if attach:
            for f in os.listdir(dt_file_path):
                path = dt_file_path + "\\" + f
                with open(path, "rb") as atch_file:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(atch_file.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition',"attachment; filename= "+f)
                    msg.attach(part)
                
        smtp.send_message(msg)
        del msg
        smtp.quit()

    def split_name_from_address(self, email_to=None):
        if email_to is None:
            email_to = self.address
        index = email_to.find("@")
        return email_to[:index]


