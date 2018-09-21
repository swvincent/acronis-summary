#! /usr/bin/env python3
"""
Acronis Backup Summary
Copyright Scott W. Vincent 2018
This program reads backup log emails from Acronis and summarizes
the information into a single email so it's easier to check on a
daily basis.
"""

from smtplib import SMTP
from email.mime.multipart import MIMEMultipart      # building new email
from email.mime.text import MIMEText                # building new email
from retrying import retry                          # retrying send email
import logging                                      # Logger
import email                                        # Email parsing
import poplib                                       # Get emails from server
from datetime import datetime                       # get current date/time
from html2text import html2text                     # plaintext email output
from more_itertools import unique_everseen          # Removing dupe errors
import dateutil.parser                              # Parse date from email
import dateutil.tz                                  # Prase date from email
# from dateutil import parser, tz                     # Parse date from email
from configparser import ConfigParser              # for config INI


def setup_logger():
    """
    Setup loggers for file and screen
    """
    global logger

    # Create logger
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s %(name)-12s '
                                  '%(levelname)-8s %(message)s')

    # Log to file
    fh = logging.FileHandler('acronsum.log')
    fh.setLevel(logging.INFO)
    fh.setFormatter(formatter)
    logger.addHandler(fh)

    # Log to screen
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)
    ch.setFormatter(formatter)
    logger.addHandler(ch)

    logger.info('Program Started')


def process_emails(mail_server, from_email, to_email, pop_user, pop_password):
    """ Process emails received and generate summary email """

    email_data = []

    m = poplib.POP3_SSL(mail_server)
    m.user(pop_user)
    m.pass_(pop_password)

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
                    # Parse date so it can be formatted, etc.
                    # http://stackoverflow.com/a/12160056
                    email_date = dateutil.parser.parse(parsed_email['Date'])
                    email_data.append([part.get_payload(), email_date])
                else:
                    logger.info('Ignoring message of type "{}"'
                                .format(part.get_content_type()))
            # Delete email
            m.dele(i+1)

        try:
            send_backups_email(mail_server, from_email, to_email, email_data)
        except Exception as ex:
            # Email failed to send after several retries. Cancel deletes.
            logger.error('Could not send backup log summary email '
                         'after several attempts: ' + str(ex))
            m.rset()
            m.quit()
        else:
            logger.info('Backup log summary email sent')
            m.quit()
    else:
        # No messages found
        m.quit()
        try:
            send_no_messages_email(mail_server, from_email, to_email)
        except Exception as ex:
            # Email failed to send after several retries
            logger.error('Could not send backup log empty email '
                         'after several attempts: ' + str(ex))
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

    html_output = ('<ul><li>' + '</li><li>'
                   .join(unique_everseen(backup_errors)) + '</li></ul>')
    return html_output


def send_backups_email(mail_server, from_email, to_email, email_data):
    """
    Send backups summary email.
    """

    msg = MIMEMultipart('alternative')
    msg['Subject'] = ('Backup Log Summary as of {:%a, %-m/%-d/%Y at %I:%M %p}'
                      .format(datetime.now()))
    msg['From'] = from_email
    msg['To'] = to_email

    # Build HTML message as unordered list
    # Based partly on http://stackoverflow.com/a/10716137
    linesUL = '<ol>'
    for edata in email_data:
        # Make sure edata[0] isn't empty string, it happened
        # once resulting in index out of range error
        if edata[0]:
            # Grab last line of email and strip period
            # filter is used to remove empty lines which
            # will cause errors with rstrip
            edata_lines = edata[0].splitlines()
            last_line = list(filter(None, edata_lines))[-1].rstrip('.')
            formatted_date = (edata[1].astimezone(dateutil.tz.tzlocal())
                              .strftime('%a, %-m/%-d/%Y at %I:%M %p'))

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

            linesUL += ('<li style="color:{}">{} on {}{}</li>'
                        .format(htmlColor, last_line,
                                formatted_date, error_info))
    linesUL += '</ol>'

    htmlMsg = '<html><head></head><body>{}</body></html>'.format(linesUL)
    htmlPart = MIMEText(htmlMsg, 'html')
    msg.attach(htmlPart)

    # Convert html to plain text and attach to message
    textMsg = html2text(htmlMsg)
    textPart = MIMEText(textMsg, 'plain')
    msg.attach(textPart)

    send_email(msg, mail_server)
    logger.debug(textMsg)


def send_no_messages_email(mail_server, from_email, to_email):
    """
    Send email that backup log is empty (in case it shouldn't be.)
    """

    msg = MIMEText('The backup log inbox is empty.')
    msg['Subject'] = ('Backup Log is empty as of {:%a, %-m/%-d/%Y at %I:%M %p}'
                      .format(datetime.now()))
    msg['From'] = from_email
    msg['To'] = to_email
    send_email(msg, mail_server)


@retry(wait_fixed=60000, stop_max_attempt_number=15)
def send_email(msg, mail_server):
    """
    Send email. Retry every minute for up to 15 minutes.
    """
    logger.debug('Attemping to send email')
    smtp = SMTP(mail_server)
    smtp.send_message(msg)
    smtp.quit()


def main():
    setup_logger()

    try:
        # Load settings from config
        config = ConfigParser()
        config.read('acronsum.ini')
        mail_server = config['main']['mail_server']
        from_email = config['main']['from_email']
        to_email = config['main']['to_email']
        pop_user = config['main']['pop_user']
        pop_password = config['main']['pop_password']
    except Exception as ex:
        logger.error('Could not load configuration: ' + str(ex))
    else:
        process_emails(mail_server, from_email, to_email,
                       pop_user, pop_password)


if __name__ == '__main__':
    main()
