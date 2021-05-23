import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate
from pathlib import Path
from string import Template

# automatic email script, originally made for McGill PTOT Orientation Week 2021 by Nicholas Yiphoiyen
# three .txt files are needed, see examples of .txt files in the repo
# 1. NAME_EMAIL.txt --> has the names and emails of the sponsors separated by a comma, one sponsor per line (I.E firstname lastname ,email)
# 2. ALREADY_SENT_EMAILS.txt --> emails, name of the sponsors that were properly sent an email
# 3. TEMPLATE_TEXT --> text that you want to send to the sponsors, set the name of the person as ${PERSON_NAME}
# you need to set the subject of the email and add any attachments you want to send with the email as an array (don't forget to put the file extension ex: .pdf, .png)
# don't forget to put your email address, password also
# Two-step verification must be disabled on your email account & allow less-secure apps to access your account (just google it)
# CAREFUL : Test it out with your personal email before sending the emails to sponsors

MY_ADDRESS = "PLACEHOLDER EMAIL"
PASSWORD = "PLACEHOLDER PASSWORD"
NAME_EMAIL = "NAME_EMAIL.txt"
ALREADY_SENT_EMAILS = "ALREADY_SENT_EMAILS.txt"
TEMPLATE_TEXT = "FROSH_EMAIL.txt"
SUBJECT = "PLACEHOLDER SUBJECT"
FILES = ["DUMMY_PDF_1", "DUMMY_PDF_2", "DUMMY_PDF_3"]

s = smtplib.SMTP(host='smtp.gmail.com', port=587)
s.starttls()
s.login(MY_ADDRESS, PASSWORD)


def get_contacts(filename):
    names = []
    emails = []
    with open(filename, mode='r', encoding='utf-8') as contacts_file:
        for a_contact in contacts_file:
            names.append(a_contact.split(', ')[0])
            emails.append(a_contact.split(', ')[1])
    return names, emails


def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)


names, emails = get_contacts(NAME_EMAIL)
message_template = read_template(TEMPLATE_TEXT)

with open(ALREADY_SENT_EMAILS, mode='a', encoding='utf-8') as already_sent_emails:
    with open(NAME_EMAIL, mode='r+', encoding='utf-8') as contacts_file:
        contacts_file.truncate(0)
        for name, email in zip(names, emails):
            try:
                msg = MIMEMultipart()

                message = message_template.substitute(PERSON_NAME=name.title())

                msg['From'] = MY_ADDRESS
                msg['To'] = email
                msg['Date'] = formatdate(localtime=True)
                msg['Subject'] = SUBJECT

                msg.attach(MIMEText(message, 'plain'))

                for path in FILES:
                    part = MIMEBase('application', "octet-stream")
                    with open(path, 'rb') as file:
                        part.set_payload(file.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition',
                                    'attachment; filename="{}"'.format(Path(path).name))
                    msg.attach(part)

                s.send_message(msg)

                print("Email sent to " + name + " at " + email)
                del msg
                already_sent_emails.write(email + " " + name + '\n')

            except Exception as e:
                print("There was an error with the email " + email + " addressed to " + name)
                contacts_file.write(email + " " + name + '\n')

print("\n\nDone sending all emails")
