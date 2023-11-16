import smtplib
from secret import pswd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# Setup port number and server name

smtp_port = 587                 #Standard secure SMTP port
smtp_server = "smtp.gmail.com"  #Google SMPT Server

email_from = 'ablickson@gmail.com'
email_list = ['ablickson@gmail.com', 'ablickson@gmail.com', 'ablickson@gmail.com']

password = pswd

subject = "Rise of the Machinea for Party Time"

def send_emails(email_list):
    
    for person in email_list:

        body = f"""
        Line 1 tank let's stomp
        Line 2 tanks let's stomp
        Line 3 tanks let's stomp
        Everybody bow to Skynet
        """

        #make a MIME object to define parts of the email

        msg = MIMEMultipart()
        msg['From'] = email_from
        msg['To'] = person
        msg['Subject'] = subject

        #Attach the body of the message
        msg.attach(MIMEText(body, 'plain'))


        # Define the file name to attach
        filename = 'DCOILWTICO.xlsx'

        # Open the file in python as a binary
        attachment = open(filename, 'rb') # r for read and b for binary

        # Encode as base 64
        attachment_package = MIMEBase('application', 'octet-stream')
        attachment_package.set_payload((attachment).read())
        encoders.encode_base64(attachment_package)
        attachment_package.add_header('Content-Disposition', "attachment; filename= " + filename)
        msg.attach(attachment_package)

        text = msg.as_string()

        # Connect with the server
        print("Connecting to server...")
        TIE_server = smtplib.SMTP(smtp_server, smtp_port)
        TIE_server.starttls()
        TIE_server.login(email_from, pswd)
        print("We connected to the server!")
        print()

        # Send emails to the ones in the list is iterated
        print(f"Sending email to: {person}...")
        TIE_server.sendmail(email_from, person, text)
        print(f"Email sent to: {person}")
        print()

    TIE_server.quit()


send_emails(email_list)
