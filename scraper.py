import sys
import getpass
import imaplib
import datetime
import csv

from email.header import decode_header
from email.parser import HeaderParser

parser = HeaderParser()

GET_EMAILS_FROM_DATE = (datetime.datetime.now() - datetime.timedelta(1)).strftime("%d-%b-%Y")
IMAP4_SSL_HOST = 'imap.gmail.com'


class ScrapeEmails:
    def __init__(self, e):
        self.get_emails_from = GET_EMAILS_FROM_DATE
        self.email_addr = e
        self.subjects = []
        self.m = imaplib.IMAP4_SSL(IMAP4_SSL_HOST)

    def login(self):
        self.m.login(self.email_addr, getpass.getpass())
        print "Login success."

    def get_data(self):
        self.m.select(mailbox='INBOX', readonly=True)
        typ, data = self.m.search(None, '(SINCE "%s")' % self.get_emails_from)
        for resp in data[0].split():
            typ2, msg_data = self.m.fetch(resp, '(RFC822)')
            header_data = msg_data[0][1]
            msg = parser.parsestr(header_data)
            self.subjects.append(msg['Subject'])
        self.subjects = list(reversed(self.subjects))

    def logout(self):
        self.m.logout()
        print "Logged out of " + self.email_addr


class FormatAndStore:
    def __init__(self):
        self.email_addr = sys.argv[1]
        self.connect = ScrapeEmails(self.email_addr)

    def get_email_subjects(self):
        self.connect.login()
        self.connect.get_data()
        self.write_subjects()

    @staticmethod
    def decode(s):
        new_str = decode_header(s)
        return new_str[0][0]

    def write_subjects(self):
        with open("subjects.txt", 'wb') as f:
            writer = csv.writer(f)
            for subject in self.connect.subjects:
                subject = self.decode(subject)
                writer.writerow([subject])
        self.logout()

    def logout(self):
        self.connect.logout()


if __name__ == '__main__':
    FormatAndStore().get_email_subjects()
