#!/usr/bin/env python3

from __future__ import print_function

import os
import re
import email
import imaplib
import hashlib
import datetime
from message import Message


class MailboxClient:

    def __init__(self, host, port, username, password):
        self.username = username
        self.mailbox = imaplib.IMAP4_SSL(host, port)
        try:
            self.mailbox.login(username, password)
        except imaplib.IMAP4.error:
            print('Unable to login to: ', username)

    def copy_emails(self, days, local_folder, wkhtmltopdf):

        n_saved = 0
        n_exists = 0

        self.local_folder = local_folder
        self.wkhtmltopdf = wkhtmltopdf
        criterion = 'ALL'

        if days:
            date = (datetime.date.today() - datetime.timedelta(days)).strftime('%d-%b-%Y')
            criterion = '(SENTSINCE {date})'.format(date=date)

        typ, data = self.mailbox.search(None, criterion)
        for num in data[0].split():
            typ, data = self.mailbox.fetch(num, '(RFC822)')
            if self.saveEmail(data):
                n_saved += 1
            else:
                n_exists += 1

        return (n_saved, n_exists)

    def cleanup(self):
        self.mailbox.close()
        self.mailbox.logout()

    def getEmailFolder(self, msg, data):
        if msg['Message-Id']:
            foldername = re.sub('[^a-zA-Z0-9_\-\.()\s]+', '', msg['Message-Id'])
        else:
            foldername = hashlib.sha224(data).hexdigest()

        year = 'None'
        if msg['Date']:
            match = re.search('\d{1,2}\s\w{3}\s(\d{4})', msg['Date'])
            if match:
                year = match.group(1)

        return os.path.join(self.local_folder, year, foldername)

    def saveEmail(self, data):
        for response_part in data:
            if isinstance(response_part, tuple):
                try:
                    # See: https://docs.python.org/3/howto/unicode.html#python-s-unicode-support
                    msg = email.message_from_string(response_part[1].decode('utf-8', 'ignore'))
                except AttributeError:
                    msg = email.message_from_string(response_part[1])

                directory = self.getEmailFolder(msg, data[0][1])

                if os.path.exists(directory):
                    return False

                os.makedirs(directory)

                try:
                    message = Message(directory, msg)
                    message.createRawFile(data[0][1])
                    message.createMetaFile()
                    message.extractAttachments()

                    if self.wkhtmltopdf:
                        message.createPdfFile(self.wkhtmltopdf)

                except Exception as e:
                    # ex: Unsupported charset on decode
                    print(directory)
                    print('MailboxClient.saveEmail() failed')
                    print(e)

        return True
