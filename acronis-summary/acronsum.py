#! /usr/bin/env python3
"""
Acronis Backup Summary
Copyright Scott W. Vincent 2018
This program reads backup log emails from Acronis and summarizes the information
into a single email so it's easier to check on a daily basis.
"""

from smtplib import SMTP
from email.mime.multipart import MIMEMultipart      # building new email
from email.mime.text import MIMEText                # building new email
from retrying import retry                          # retrying if email send fails
import logging                                      # Logger
import email                                        # Email parsing
import poplib                                       # Get emails from server
from datetime import datetime                       # get current date/time
from html2text import html2text                     # produce plaintext email output
from more_itertools import unique_everseen          # Removing dupes from error list
import dateutil.parser                              # Parse date from email
import dateutil.tz                                  # Prase date from email


def setup_logger():
    """
    Setup loggers for file and screen
    """
    global logger

    # Create logger
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s %(name)-12s %(levelname)-8s %(message)s')

    Log to file
    fh = logging.FileHandler('not.log')
    fh.setLevel(logging.INFO)
    fh.setFormatter(formatter)
    logger.addHandler(fh)

    # Log to screen
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)
    ch.setFormatter(formatter)
    logger.addHandler(ch)

    logger.info('Program Started')


def process_emails():
    email_data = []

    m = poplib.POP3_SSL(MAIL_SERVER)
    m.user(POP_USER)
    m.pass_(POP_PASS)

    num_messages = len(m.list()[1])

    if num_messages > 0:
        # New messages found, process them
        for i in range(num_messages):
            raw_email = b"\n".join(m.retr(i+1)[1])
            parsed_email = email.message_from_bytes(raw_email)
            # Acronis emails are just plain text, but other emails may come to
            # inbox for account so I verify that it's plain text to be safe.
            for part in parsed_email.walk():
                if part.get_content_type() == 'text/plain':
                    # Parse date so it can be formatted, etc. http://stackoverflow.com/a/12160056
                    email_date = dateutil.parser.parse(parsed_email['Date'])
                    email_data.append([part.get_payload(), email_date])
                else:
                    logger.info('Ignoring message of type "{}"'.format(part.get_content_type()))
            # Delete email
            m.dele(i+1)

        try:
            send_backups_email(email_data)
        except Exception as ex:
            # Email failed to send after several retries. Cancel deletes.
            logger.error('Could not send backup log summary email after several attempts: ' + str(ex))
            m.rset()
            m.quit()
        else:
            logger.info('Backup log summary email sent')
            m.quit()
    else:
        # No messages found
        m.quit()
        try:
            send_no_messages_email()
        except Exception as ex:
            # Email failed to send after several retries
            logger.error('Could not send backup log empty email after several attempts: ' + str(ex))
        else:
            logger.info('Backup log empty email sent')


def extract_errors(email_text):
    """
    Extract error information from Acronis email
    and format in HTML Unordered List
    """
    email_lines = email_text.splitlines()
    backup_errors = []

    for email_line in email_lines:
        if email_line.startswith('Error code:'):
            backup_error = email_line
        elif email_line.startswith('Message:'):
            backup_error += ':{}'.format(email_line[8:])
            backup_errors.append(backup_error)

    html_output = '<ul><li>' + '</li><li>'.join(unique_everseen(backup_errors)) + '</li></ul>'
    return html_output


def send_backups_email(email_data):
    """
    Send backups summary email.
    """

    msg = MIMEMultipart('alternative')
    msg['Subject'] = 'Backup Log Summary as of {:%a, %-m/%-d/%Y at %I:%M %p}'.format(datetime.now())
    msg['From'] = FROM_EMAIL
    msg['To'] = TO_EMAIL

    # Build HTML message as unordered list
    # Based partly on http://stackoverflow.com/a/10716137
    linesUL = '<ol>'
    for edata in email_data:
        # Make sure edata[0] isn't empty string, it happened
        # once resulting in index out of range error
        if edata[0]:
            # Grab last line of email and strip period
            # filter is used to remove empty lines which will cause errors with rstrip
            edata_lines = edata[0].splitlines()
            last_line = list(filter(None, edata_lines))[-1].rstrip('.')
            formatted_date = edata[1].astimezone(dateutil.tz.tzlocal()).strftime('%a, %-m/%-d/%Y at %I:%M %p')

            if 'has succeeded' in last_line:
                # Success, dark green
                htmlColor = '#006400'
                error_info = ""
            elif 'has failed' in last_line:
                # Failed, Red
                htmlColor = '#FF0000'
                # Add error info
                error_info = extract_errors(edata[0])
            else:
                # ??? Black
                htmlColor = '#000000'
                error_info = ""

            linesUL += '<li style="color:{}">{} on {}{}</li>'.format(htmlColor, last_line,
                                                                     formatted_date, error_info)
    linesUL += '</ol>'

    htmlMsg = '<html><head></head><body>{}</body></html>'.format(linesUL)
    htmlPart = MIMEText(htmlMsg, 'html')
    msg.attach(htmlPart)

    # Convert html to plain text and attach to message
    textMsg = html2text(htmlMsg)
    textPart = MIMEText(textMsg, 'plain')
    msg.attach(textPart)

    send_email(msg)
    logger.debug(textMsg)


def send_no_messages_email():
    """
    Send email that backup log is empty (in case it shouldn't be.)
    """

    msg = MIMEText('The backup log inbox is empty.')
    msg['Subject'] = 'Backup Log is empty as of {:%a, %-m/%-d/%Y at %I:%M %p}'.format(datetime.now())
    msg['From'] = FROM_EMAIL
    msg['To'] = TO_EMAIL
    send_email(msg)


@retry(wait_fixed=60000, stop_max_attempt_number=15)
def send_email(msg):
    """
    Send email. Retry every minute for up to 15 minutes.
    """
    logger.debug('Attemping to send email')
    smtp = SMTP(MAIL_SERVER)
    smtp.send_message(msg)
    smtp.quit()


# Global constants
MAIL_SERVER = 'mailserver.com'
POP_USER = 'username'
POP_PASS = 'password'
FROM_EMAIL = 'username@mailserver.com'
TO_EMAIL = 'someoneelse@mailserver.com'

# Entry point
setup_logger()
process_emails()