import os
import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# configure the subject of the email
month = "08"
year = "2024"

# choose subject line:
subject = f"Originals Salarisspecificatie {month}-{year}"
# subject = f"Originals Urenoverzicht {month}-{year}"

# read the config file
with open('config.json', 'r') as f:
    data = json.load(f)

# Extract the contacts and email credentials from the json data
config = data['contacts']
host = data['host']

# directory containing the PDF files
pdf_folder = './documents'

# email settings
smtp_server = 'smtp-mail.outlook.com'
smtp_port = 587

# establish a secure session with outlook's outgoing SMTP server using your outlook account
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()
server.login(host['email'], host['passw'])


# loop through each person and email in the config
for pdf, person in config.items():
    pdf_path = os.path.join(pdf_folder, f'{pdf}.pdf')
    if not os.path.exists(pdf_path):
        print(f"File {pdf_path} does not exist, skipping.")
        continue

    # set up the email
    msg = MIMEMultipart()
    msg['From'] = host['email']
    msg['To'] = person
    msg['Subject'] = subject

    # attach the pdf
    with open(pdf_path, 'rb') as f:
        attach = MIMEBase('application', 'octet-stream')
        attach.set_payload(f.read())
    encoders.encode_base64(attach)
    attach.add_header('Content-Disposition', 'attachment', filename=f"{pdf.capitalize()} {msg['Subject']}.pdf")
    msg.attach(attach)

    # add footer text
    footer_text = "This is an automated message. Please do not reply."
    msg.attach(MIMEText(footer_text, 'plain'))

    # send the email
    server.send_message(msg)

    # print the success message
    print(f"Successfully sent {pdf} to {person}")

    # remove the file from the directory
    os.remove(pdf_path)
    print(f"Removed {pdf_path}\n")

# close the connection to the server
server.quit()
