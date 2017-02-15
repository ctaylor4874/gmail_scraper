import sys
import getpass
import imaplib
import datetime

from email.parser import HeaderParser

m = imaplib.IMAP4_SSL('imap.gmail.com')
parser = HeaderParser()


class ScrapeEmails:
    def __init__(self, e):
        self.get_emails_from = (datetime.datetime.now() - datetime.timedelta(1)).strftime("%d-%b-%Y")
        self.email_addr = e
        self.m = imaplib.IMAP4_SSL('imap.gmail.com')
        try:
            self.m.login(self.email_addr, getpass.getpass())
            print("Logged in as " + self.email_addr)

        except imaplib.IMAP4.error:
            print "Login Failed! Check username and password"

    def get_data(self):
        self.m.select(mailbox='INBOX')
        typ, data = self.m.search(None, '(SINCE "%s")' % self.get_emails_from)
        for resp in data[0].split():
            typ2, msg_data = self.m.fetch(resp, '(RFC822)')
            header_data = msg_data[0][1]
            msg = parser.parsestr(header_data)
            print msg['Subject']

    def logout(self):
        self.m.logout()
        print "Logged out of " + self.email_addr


email_addr = sys.argv[1]
connect = ScrapeEmails(email_addr)
connect.get_data()
# todo store data in a file
connect.logout()
