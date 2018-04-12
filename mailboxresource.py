#!/usr/bin/env python3

import os
import re
import email
import hashlib
import logging
import imaplib
from email import policy
from message import Message
from datetime import date, timedelta
from imapclient import IMAPClient

logging.basicConfig(
    filename='imapbox.log',
    format='%(asctime)s - %(levelname)s: %(message)s',
    datefmt='%Y-%m-%dТ%H:%M:%S%z',
    level=logging.INFO
)


def split(arr, size):
    arrs = []
    while len(arr) > size:
        pice = arr[:size]
        arrs.append(pice)
        arr = arr[size:]
    arrs.append(arr)
    return arrs


class MailboxClient:

    def __init__(self, name, host, port, username, password, remote_folder):
        self.name = name
        self.host = host
        self.port = port
        self.username = username
        self.password = password
        self.remote_folder = remote_folder

        self.mailbox = IMAPClient(self.host, self.port)

        try:
            self.mailbox.login(self.username, self.password)
        except imaplib.IMAP4.error:
            logging.exception('Unable to login to: %s', self.username)

    def copy_emails(self, days, limit, local_folder):
        self.days = days
        self.limit = limit
        self.local_folder = local_folder
        self.saved = 0
        self.total = 0

        criterion = 'ALL'

        if self.days:
            since = date.today() - timedelta(self.days)
            criterion = ['SINCE', since.strftime('%d-%b-%Y')]

        if self.remote_folder == 'ALL':
            for flags, delimiter, folder in self.mailbox.list_folders():
                self.fetch_emails(folder, criterion)
        else:
            self.fetch_emails(self.remote_folder, criterion)

        return (self.saved, self.total)

    def fetch_emails(self, folder, criterion):
        saved = 0
        total = 0

        self.mailbox.select_folder(folder, readonly=True)

        messages = self.mailbox.search(criterion)
        for part in split(messages, self.limit):
            for msgid, data in self.mailbox.fetch(part, ['RFC822']).items():
                if self.save_email(data):
                    saved += 1

        total = len(messages)
        logging.info(
            '[%s/%s] - saved: %s, total: %s;',
            self.username,
            folder,
            saved,
            total
        )

        self.saved += saved
        self.total += total

    def logout(self):
        self.mailbox.logout()

    def get_email_folder(self, message, body):
        if message['Message-Id']:
            exp = '[^a-zA-Z0-9_\-\.\s]+'
            directory = re.sub(exp, '', message['Message-Id'])
            directory = directory.strip()
        else:
            directory = hashlib.sha3_256(body).hexdigest()

        year = 'None'
        if message['Date']:
            match = email.utils.parsedate(message['Date'])
            if match:
                year = str(match[0])

        return os.path.join(self.local_folder, year, directory)

    def save_email(self, data):
        body = data[b'RFC822']

        try:
            message = email.message_from_bytes(body, policy=policy.default)
        except Exception as e:
            logging.error(e)

        directory = self.get_email_folder(message, body)

        try:
            os.makedirs(directory)
        except FileExistsError:
            return False

        try:
            msg = Message(directory, message)
            msg.create_raw_file(body)
            msg.create_meta_file()
            msg.extract_attachments()

        except Exception as e:
            logging.info('Faulty email: ', directory)
            logging.exception(e)

        return True
