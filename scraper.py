import sys
import getpass
import imaplib
import datetime
import csv

from email.header import decode_header
from email.parser import HeaderParser

# Date to scrape emails from. must be in GMail format ("%d-%b-%Y").
GET_EMAILS_FROM_DATE = (datetime.datetime.now() - datetime.timedelta(1)).strftime("%d-%b-%Y")
# Host of the email account.
IMAP4_SSL_HOST = 'imap.gmail.com'

parser = HeaderParser()


class ScrapeEmails:
    """
    Class to scrape emails for subjects.

    Class takes in an email address provided by the user and sets up a connection via imaplib. Stores data as a list
    in self.subjects.
    """

    def __init__(self, e):
        """
        Initializiation of date to get emails from, email address, subjects, and email connector.

        :param e: Email address to get subjects from.
        :type e: str
        """
        self.get_emails_from = GET_EMAILS_FROM_DATE
        self.email_addr = e
        self.subjects = []
        self.m = imaplib.IMAP4_SSL(IMAP4_SSL_HOST)

    def login(self):
        """
        Logs the user into their email account with the email address and password provided.

        If login is successful, print Login Success.

        :return:
        """
        self.m.login(self.email_addr, getpass.getpass())
        print "Logged in as {}.".format(self.email_addr)

    def get_data(self):
        """
        Gets the data from the mailbox and organizes the data list.

        Selects inbox and sets readonly=True so the emails are not marked as read.  Searches the mailbox for
        emails received since the get_emails_from date. Fetches and parses the data and stores the 'Subject' from
        each email in self.subjects. The subjects list is then reversed to put the list in ascending order by date.

        :return:
        """
        self.m.select(mailbox='INBOX', readonly=True)
        typ, data = self.m.search(None, '(SINCE "%s")' % self.get_emails_from)
        for resp in data[0].split():
            typ2, msg_data = self.m.fetch(resp, '(RFC822)')
            header_data = msg_data[0][1]
            msg = parser.parsestr(header_data)
            self.subjects.append(msg['Subject'])
        self.subjects = list(reversed(self.subjects))

    def logout(self):
        """
        Terminates the connection to the users email account.

        :return:
        """
        self.m.logout()
        print "Logged out of {}.".format(self.email_addr)


class FormatAndStore:
    """
    Formats and stores the email data.
    """

    def __init__(self):
        """
        Gets email address from user and creates an instance of ScrapeEmails.
        """
        self.email_addr = sys.argv[1]
        self.connect = ScrapeEmails(self.email_addr)

    def get_email_subjects(self):
        """
        Calls the methods to login, get data, and write subjects to a location.

        :return:
        """
        self.connect.login()
        self.connect.get_data()
        self.write_subjects()

    @staticmethod
    def decode(s):
        """
        Decode the subjects that have been scraped to make sure they are human readable.

        :param s: Subject to decode.
        :type s: str
        :return: Human readable subject string.
        :rtype: str
        """
        new_str = decode_header(s)
        return new_str[0][0]

    def write_subjects(self):
        """
        Writes the subjects to the a file. Disconnects from users account after file is written.

        Iterates through subject list and decodes each subject.  Writes the subjects to a file.

        :return:
        """
        with open("subjects.txt", 'wb') as f:
            writer = csv.writer(f)
            for subject in self.connect.subjects:
                subject = self.decode(subject)
                writer.writerow([subject])

    def logout(self):
        """
        Makes the call to terminate the users connection to their email account.

        :return:
        """
        self.connect.logout()


if __name__ == '__main__':
    conn = FormatAndStore()
    conn.get_email_subjects()
    conn.logout()
