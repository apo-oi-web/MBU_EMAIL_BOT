
# MBU Email Bot
# By Daniel McDonough 10/15/2019

import smtplib
import pandas
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from string import Template

# Global vars
USERNAME = "USERNAME"  # Remember to change
PASSWORD = "PASSWORD"  # Remember to change
PORT = 587
SSL = "YES"
CSV_FILE = "./Badge_assignments.csv"
emailTemplate = 'message.txt'
SUBJECT = 'MBU Badge Assignments Confirmation'

# NOTE: SERVER NAME MAY CHANGE OVER TIME
# TO CHECK THE NEW SERVER NAME EXECUTE 'ping smtp.office365.com'
SERVER = "lyh-efz.ms-acdc.office.com"


# SETUP SERVER
server = smtplib.SMTP(SERVER, PORT)
server.ehlo()
server.starttls()

# Next, log in to the server
server.login(USERNAME, PASSWORD)


# Read CSV with email, name and Badge information
def readCSV():
    csv = pandas.read_csv(CSV_FILE)
    rows = csv.shape[0]

    for i in range(rows):
        email = csv["email"][i]
        name = csv["Name"][i]
        period1 = csv["1st Period"][i]
        period2 = csv["2nd Period"][i]
        period3 = csv["3rd Period"][i]
        print("Generating email for: " + str(email))

        message = genMessage(name, period1, period2, period3)
        sendEmail(email,message)


# Read the template file into memory
def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)


# Fill in the template with data needed
def genMessage(name,p1,p2,p3):
    message_template = read_template(emailTemplate)
    message = message_template.substitute(PERSON_NAME=name,PERIOD1=p1,PERIOD2=p2,PERIOD3=p3)
    return message


def sendEmail(target,message):

    msg = MIMEMultipart()

    msg['Subject'] = SUBJECT
    msg['From'] = USERNAME  # the sender's email address
    msg['To'] = target  # the recipient's email address

    # add in the message body
    msg.attach(MIMEText(message, 'plain'))

    # Send the mail
    server.sendmail(USERNAME, target, msg.as_string())


if __name__ == '__main__':
    readCSV()