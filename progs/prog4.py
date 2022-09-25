import smtplib
import os, os.path
from email import encoders
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from konfs.konfs import sender, password, getter


def prog4(pathtofile):
    """
    Отправка письма с вложением
    """
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    try:
        server.login(sender, password)
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["Subject"] = f"Журнал и акты НШЛ {os.path.splitext(os.path.basename(pathtofile))[0]}"
        msg.attach(MIMEText(f'архив во вложении'))
        with open(pathtofile, "rb") as f:
            file = MIMEBase('application', 'rar')
            file.set_payload(f.read())
            encoders.encode_base64(file)
            file.add_header('content-disposition', 'attachment', filename=f'Журнал и акты НШЛ {os.path.basename(pathtofile)}')
            msg.attach(file)
        print("Отправка...")
        server.sendmail(sender, getter, msg.as_string())
        return "Сообщение отправлено!"
    except Exception as _ex:
        return f"{_ex}\n Error!"