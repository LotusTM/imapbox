#!/usr/bin/env python3

import os
import re
import html
import json
import gzip
import time
import email
import chardet
import mimetypes
from html.parser import HTMLParser


class MLStripper(HTMLParser):
    def __init__(self):
        self.reset()
        self.fed = []

    def convert_charrefs(x):
        return x

    def handle_data(self, d):
        self.fed.append(d)

    def get_data(self):
        return ''.join(self.fed)


def strip_tags(html):
    s = MLStripper()
    s.feed(html)
    return s.get_data()


class Message:

    def __init__(self, directory, msg):
        self.msg = msg
        self.directory = directory

    def normalizeDate(self, datestr):
        t = email.utils.parsedate_to_datetime(datestr)
        rfc2822 = email.utils.format_datetime(t)
        # TODO: convert it to UTC
        iso8601 = t.isoformat()

        return (rfc2822, iso8601)

    def create_meta_file(self):
        parts = self.get_parts()
        attachments = []
        content = ''

        for file in parts['files']:
            attachments.append(file[1])

        if parts['text']:
            content = self.getTextContent(parts['text'])
        elif parts['html']:
            content = strip_tags(self.getHtmlContent(parts['html']))

        # Filter out Calendar items, they do not have a Date property
        if self.msg['Date'] is not None:
            Id = self.msg['Message-Id'].strip() if self.msg['Message-Id'] is not None else ''
            Subject = self.msg['Subject'].strip() if self.msg['Subject'] is not None else ''
            rfc2822, iso8601 = self.normalizeDate(self.msg['Date'])

            with open(os.path.join(self.directory, 'metadata.json'), 'w') as fp:
                fp.write(json.dumps({
                    'Id': Id,
                    'Subject': Subject,
                    'From': self.msg['From'],
                    'To': self.msg['To'],
                    'Cc': self.msg['Cc'],
                    'Date': rfc2822,
                    'Utc': iso8601,
                    'Attachments': attachments,
                    'WithHtml': len(parts['html']) > 0,
                    'WithText': len(parts['text']) > 0,
                    'Body': content
                }, indent=4, ensure_ascii=False))

    def create_raw_file(self, data):
        with gzip.open(os.path.join(self.directory, 'raw.eml.gz'), 'wb') as fp:
            fp.write(data)

    def getTextContent(self, parts):
        if not hasattr(self, 'text_content'):
            self.text_content = ''
            for part in parts:
                self.text_content += part.get_content()
        return self.text_content

    def createTextFile(self, parts):
        content = self.getTextContent(parts)
        with open(os.path.join(self.directory, 'message.txt'), 'w') as fp:
            fp.write(content)

    def getHtmlContent(self, parts):
        if not hasattr(self, 'html_content'):
            self.html_content = ''

            for part in parts:
                self.html_content += part.get_content()

            m = re.search('<body[^>]*>(.+)<\/body>', self.html_content, re.S | re.I)
            if (m is not None):
                self.html_content = m.group(1)

        return self.html_content

    def createHtmlFile(self, parts, embed):
        content = self.getHtmlContent(parts)
        for img in embed:
            pattern = 'src=["\']cid:%s["\']' % (re.escape(img[0]))
            path = os.path.join('attachments', img[1])
            content = re.sub(pattern, 'src="%s"' % (path), content, 0, re.S | re.I)

        subject = self.msg['Subject']
        fromname = self.msg['From']

        content = """<!doctype html>
<html>
<head>
    <meta charset="utf-8">
    <meta http-equiv="x-ua-compatible" content="ie=edge">
    <meta name="author" content="%s">
    <title>%s</title>
</head>

<body>
%s
</body>
</html>""" % (html.escape(fromname), html.escape(subject), content)

        with open(os.path.join(self.directory, 'message.html'), 'w') as fp:
            fp.write(content)

    def sanitizeFilename(self, filename):
        keepcharacters = (' ', '.', '_', '-')
        return "".join(c for c in filename if c.isalnum() or c in keepcharacters).rstrip()

    def get_parts(self):
        if not hasattr(self, 'message_parts'):
            counter = 1
            message_parts = {
                'text': [],
                'html': [],
                'embed_images': [],
                'files': []
            }

            for part in self.msg.walk():
                # multipart/* are just containers
                if part.get_content_maintype() == 'multipart':
                    continue

                filename = part.get_filename()
                if not filename:
                    if part.get_content_type() == 'text/plain':
                        message_parts['text'].append(part)
                        continue

                    if part.get_content_type() == 'text/html':
                        message_parts['html'].append(part)
                        continue

                    ext = mimetypes.guess_extension(part.get_content_type())
                    if not ext:
                        # Use a generic bag-of-bits extension
                        ext = '.bin'
                    filename = 'part-%03d%s' % (counter, ext)

                filename = self.sanitizeFilename(filename)

                content_id = part.get('Content-Id')
                if (content_id):
                    content_id = content_id[1:][:-1]
                    message_parts['embed_images'].append((content_id, filename))

                counter += 1
                message_parts['files'].append((part, filename))
            self.message_parts = message_parts
        return self.message_parts

    def extract_attachments(self):
        message_parts = self.get_parts()

        if message_parts['text']:
            self.createTextFile(message_parts['text'])

        if message_parts['html']:
            self.createHtmlFile(message_parts['html'], message_parts['embed_images'])

        if message_parts['files']:
            directory = os.path.join(self.directory, 'attachments')

            try:
                os.makedirs(directory)
            except FileExistsError:
                pass

            for file in message_parts['files']:
                with open(os.path.join(directory, file[1]), 'wb') as fp:
                    payload = file[0].get_payload(decode=True)
                    if payload:
                        fp.write(payload)
