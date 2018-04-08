#!/usr/bin/env python3

import os
import re
import email
import hashlib
import logging
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
            logging.exception('Unable to login to: ', self.username)

    def copy_emails(self, days, limit, local_folder):
        self.days = days
        self.limit = limit
        self.local_folder = local_folder
        self.saved = 0

        criterion = 'ALL'

        if self.days:
            since = date.today() - timedelta(self.days)
            criterion = ['SINCE', since.strftime('%d-%b-%Y')]

        if self.remote_folder == 'ALL':
            for flags, delimiter, folder in self.mailbox.list_folders():
                self.fetch_emails(folder, criterion)
        else:
            self.fetch_emails(self.remote_folder, criterion)

        return self.saved

    def fetch_emails(self, folder, criterion):
        n_saved = 0

        self.mailbox.select_folder(folder, readonly=True)

        messages = self.mailbox.search(criterion)
        for part in split(messages, self.limit):
            for msgid, data in self.mailbox.fetch(part, ['RFC822']).items():
                if self.save_email(data):
                    n_saved += 1

        logging.info('[%s/%s] - saved: %s;', self.username, folder, n_saved)

        self.saved += n_saved

    def logout(self):
        self.mailbox.logout()

    def get_email_folder(self, message, body):
        if message['Message-Id']:
            exp = '[^a-zA-Z0-9_\-\.\s]+'
            foldername = re.sub(exp, '', message['Message-Id'])
            foldername = foldername.strip()
        else:
            foldername = hashlib.sha3_256(body).hexdigest()

        year = 'None'
        if message['Date']:
            # TODO: replace with email.utils.parsedate()
            match = re.search('\d{1,2}\s\w{3}\s(\d{4})', message['Date'])
            if match:
                year = match.group(1)

        return os.path.join(self.local_folder, year, foldername)

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
            msg.createMetaFile()
            msg.extract_attachments()

        except Exception as e:
            logging.error('Faulty email: ', directory)
            logging.exception(e)

        return True
