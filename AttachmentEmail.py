
# MBU Email Bot
# By Daniel McDonough 11/3/2019
# This version includes the ability to have attachements and HTML text editiing

import smtplib
import pandas
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from string import Template
from email.mime.application import MIMEApplication


# Global vars
USERNAME = "USERNAME"  # Remember to change
PASSWORD = "PASSWORD"  # Remember to change
PORT = 587
SSL = "YES"
CSV_FILE = "./Badge_assignments.csv"
emailTemplate = './addendum.txt'
SUBJECT = 'MBU Lunch and Parking Information [IMPORTANT]'
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
    # message = message_template
    return message

def basename(filename):
    return filename.split("/")[2]

def sendEmail(target,message):

    msg = MIMEMultipart()

    msg['Subject'] = SUBJECT
    msg['From'] = USERNAME  # the sender's email address
    msg['To'] = target  # the recipient's email address

    inner = MIMEText(message, 'html')

    # files = ["./Attachments/Getting_to_Campus.pdf","./Attachments/MBU_Dropoff_and_Pickup.pdf","./Attachments/MBU_schedule.pdf","./Attachments/Minors_on_Campus.pdf", "./Attachments/Things_to_do_in_Worcester.pdf"]
    files = ["./Attachments/Minors_on_Campus.pdf",
             "./Attachments/ParkingPass.pdf"]

    for f in files:
        with open(f, "rb") as fil:
            part = MIMEApplication(
                fil.read(),
                Name=basename(f)
            )
        # After the file is closed
        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
        msg.attach(part)

    # add in the message body
    msg.attach(inner)

    # Send the mail
    server.sendmail(USERNAME, target, msg.as_string())


if __name__ == '__main__':
    readCSV()