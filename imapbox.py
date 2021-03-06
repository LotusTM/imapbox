#!/usr/bin/env python3

import os
import logging
import argparse
import configparser
from mailboxresource import MailboxClient

logging.basicConfig(
    filename='imapbox.log',
    format='%(asctime)s - %(levelname)s: %(message)s',
    datefmt='%Y-%m-%dТ%H:%M:%S%z',
    level=logging.INFO,
)


def load_configuration(args):
    config = configparser.ConfigParser(allow_no_value=True)
    config.read(os.path.expanduser(args.config_path))

    options = {
        'days': config.get('imapbox', 'days', fallback=args.days),
        'limit': config.get('imapbox', 'limit', fallback=args.limit),
        'local_folder': os.path.expanduser(
            config.get('imapbox', 'local_folder', fallback=args.local_folder)
        ),
        'accounts': []
    }

    # TODO: will not work with empty config and credentials passed as args
    for section_name in config.sections():

        if (section_name == 'imapbox'):
            continue

        if (args.account and (args.account != section_name)):
            continue

        section = config[section_name]
        account = {
            'name': section_name,
            'host': section.get('host', args.host),
            'port': section.get('port', args.port),
            'username': section.get('username', args.username),
            'password': section.get('password', args.password),
            'remote_folder': section.get('remote_folder', args.remote_folder)
        }

        if ((account['host'] or account['username'] or account['password']) is None):
            continue

        options['accounts'].append(account)

    return options


def main():
    argparser = argparse.ArgumentParser(
        description='Export messages into .eml files using IMAP protocol',
    )
    argparser.add_argument(
        '-host',
        dest='host',
        help='IMAP server host name',
    )
    argparser.add_argument(
        '-port',
        dest='port',
        help='IMAP server port',
        type=int,
        default=993,
    )
    argparser.add_argument(
        '-u',
        dest='username',
        help='Username to access email account',
    )
    argparser.add_argument(
        '-p',
        dest='password',
        help='Password to access email account',
    )
    argparser.add_argument(
        '-c',
        dest='config_path',
        help='Path to configuration file',
        default='config.ini',
    )
    argparser.add_argument(
        '-l',
        dest='local_folder',
        help='Local folder where to dump emails',
        default='./archive',
    )
    argparser.add_argument(
        '-r',
        dest='remote_folder',
        help='Remote IMAP folder that should be backed up',
        default='INBOX',
    )
    argparser.add_argument(
        '-d',
        dest='days',
        help='How many days of correspondence to back up',
        type=int,
    )
    argparser.add_argument(
        '-limit',
        dest='limit',
        help='How many emails to get during single fetch',
        type=int,
        default=10,
    )
    argparser.add_argument(
        '-a',
        dest='account',
        help='Select a specific account to backup',
    )
    args = argparser.parse_args()
    options = load_configuration(args)

    for account in options['accounts']:

        logging.info(
            '[%s/%s] start email fetching from %s:%s;',
            account['name'],
            account['remote_folder'],
            account['host'],
            account['port'],
        )

        mailbox = MailboxClient(**account)
        stats = mailbox.copy_emails(
            options['days'],
            options['limit'],
            options['local_folder']
        )
        mailbox.logout()

        logging.info(
            '[%s/%s] downloaded emails %s/%s;',
            account['name'],
            account['remote_folder'],
            stats[0],
            stats[1],
        )


if __name__ == '__main__':
    main()
