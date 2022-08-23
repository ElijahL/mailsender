import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

class MailSender():
    
    def __init__(self, username, password, smtp_server, port):
        self.smtp_server = smtp_server
        self.port = port
        self.username = username
        self.password = password

    def send_mail(self, sender, receiver, cc, subject, body, attachment_paths=None):

        message = MIMEMultipart()
        message["From"] = sender
        message["To"] = receiver
        message["Subject"] = subject
        message["CC"] = cc
        receiver += ',' + cc
        message.attach(MIMEText(body, "plain"))

        if attachment_paths:
            for path in attachment_paths:
                with open(path, 'rb') as file:
                    part = MIMEApplication(
                        file.read(),
                        Name=basename(path)
                    )
                part['Content-Disposition'] = 'attachment; filename="%s"' % basename(path)
                message.attach(part)

        #send your message with credentials specified above
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()

            server.login(self.username, self.password)
            server.sendmail(sender, receiver, message.as_string())
