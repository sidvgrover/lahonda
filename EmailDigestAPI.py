import os
import xlwt
import smtplib
import email.utils
import numpy as np
from os.path import basename
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import COMMASPACE, formatdate
from email.mime.application import MIMEApplication

class EmailDigestAPI(object):
    def __init__(self, username, password):
        self.server = smtplib.SMTP('smtp.gmail.com:587')
        self.server.starttls()
        self.username = username
        self.password = password
        self.server.login(username, password)
        print 'Server successfully initialized'


    def send_mail(self, recipient_email, subject, text, files = None):
        msg = MIMEMultipart()
        msg['To'] = email.utils.formataddr(('Recipient', recipient_email))
        msg['From'] = email.utils.formataddr(('La Honda Daily Digest', self.username))
        msg['Subject'] = subject
        msg['Date'] = formatdate(localtime = True)
        msg.attach(MIMEText(text))

        print 'Created message skeleton'

        for f in files or []:
            with open(f, "rb") as fil:
                msg.attach(MIMEApplication(fil.read(), Content_Disposition='attachment; filename="%s"' % basename(f), Name=basename(f)))

        print 'Attached all the files'
        self.server.sendmail(self.username, recipient_email, msg.as_string())

    # master_file_name is the relevant SEC-recent CSV for the day
    # user_email is the user to whom to send the keyword-modified master_file_name

    def quit(self):
        print "Server closed"
        self.server.close()