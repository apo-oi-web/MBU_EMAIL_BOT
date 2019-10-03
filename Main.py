
import smtplib
import pandas
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from string import Template


# NOTE: SERVER NAME MAY CHANGE OVER TIME
# TO CHECK THE NEW SERVER NAME EXECUTE 'ping smtp.office365.com'
SERVER = "lyh-efz.ms-acdc.office.com"
USERNAME = "OUTLOOKEMAIL"
PASSWORD = "PASSWORD"
PORT = 587
SSL = "YES"

#SETUP SERVER
server = smtplib.SMTP(SERVER, PORT)
server.ehlo()
server.starttls()

# Next, log in to the server
server.login(USERNAME, PASSWORD)


def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)


def readCSV():
    csv = pandas.read_csv("./Badge_assignments.csv")
    rows = csv.shape[0]

    for i in range(rows):
        email = csv["email"][i]
        name = csv["Name"][i]
        period1 = csv["1st Period"][i]
        period2 = csv["2nd Period"][i]
        period3 = csv["3rd Period"][i]
        print(email)
        print(name)
        print(period1)
        print(period2)
        print(period3)

        print("\n\n\n")

        message = genMessage(name, period1, period2, period3)
        sendEmail(email,message)



def genMessage(name,p1,p2,p3):

    message_template = read_template('message.txt')
    message = message_template.substitute(PERSON_NAME=name,PERIOD_ONE=p1,PERIOD2=p2,PERIOD3=p3)
    print(message)
    print(type(message))
    return message



def sendEmail(target,message):
    # Send the mail

    msg = MIMEMultipart()
    # me == the sender's email address
    # you == the recipient's email address
    msg['Subject'] = 'MBU Badge Assignments [TEST]'
    msg['From'] = USERNAME
    msg['To'] = target


    # add in the message body
    msg.attach(MIMEText(message, 'plain'))

    server.sendmail(USERNAME, target, msg.as_string())





def main():
    # Send the mail
    msg = "Hello!"  # The /n separates the message from the headers
    server.sendmail(USERNAME, USERNAME, msg)


if __name__ == '__main__':
    # main()
    readCSV()