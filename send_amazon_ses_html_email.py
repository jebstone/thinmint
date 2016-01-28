#!/usr/bin/python
import time
import pyodbc
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
 
thinmint_db = {
    "Driver": "{SQL Server Native Client 11.0}",
    "Server": "NEPTUNE\SQLEXPRESS",
    "Integrated Security": "True",
    "Trusted_Connection": "yes",
    "Database": "ThinMint"
}
 
#Change according to your settings
smtp_server = 'email-smtp.us-east-1.amazonaws.com'
smtp_username = ''
smtp_password = ''
smtp_port = '587'
smtp_do_tls = True
 
ses_send_rate = 1 / 14
 
# Email Config Info
html_file_location = "..\ThinMint\Email\MothersDay2015\ThinMint_MothersDay2015.htm"
#subject_lines = ["Mothers' Day Gifts She'll Love","Last-Minute Mothers' Day Gifts"]
 
 
email_from = "defunct_address@thinmint.io"

 
# Read the HTML email into a variable
f = open(html_file_location)
html = f.read()
 
 
def get_user_list():
""" Retrieves list of top N UserIDs and Email Addresses from list that haven't previously received an email"""

    db_connection = pyodbc.connect(**thinmint_db)
    db_cursor = db_connection.cursor()
 
    db_cursor.execute("select top 5000 u.userID, u.email from dbo.Users u left join dbo.EmailSends s on u.userid = s.userid where s.userid is null and u.active = 1")
    #db_cursor.execute("select top 10 u.userID, u.email from dbo.Users u left join dbo.EmailSends s on u.userid = s.userid where s.userid is null and u.active = 1 and u.LeadSource = 'seed'")
    users = db_cursor.fetchall()
    db_connection.close()
 
    # For Testing
    # users = []
    # users.append([10000,"me@somewhere.com"])
 
    return(users)
 
 
 
def log_email_send(user):
""" Logs email send data to a tracking database """

    userid = user[0]
    email_to = user[1]
    campaign = "MothersDay2015"
    # Set up database connection
    db_connection = pyodbc.connect(**thinmint_db)
    db_cursor = db_connection.cursor()
 
    # Load all new signups to database
    db_cursor.execute("execute dbo.LogEmailSend ?,?,?", userid, email_to, campaign)
    db_cursor.commit()
    db_connection.close()
 
 
def update_user_comm_status(userid):
    pass


def customize_html(user):
""" Accepts a list of userID and email address.
Customizes an HTML email template with userdata and returns HTML
"""
    userid = user[0]
    email_to = user[1]
    custom_html = html.replace("{{userid}}", str(userid))
    return custom_html
 
     
def generate_mime_message(user):
"""Accepts a list of userID and email address.
Generates a multi-part MIME email message and returns it
"""

    userid = user[0]
    email_to = user[1]
 
    mime_message = MIMEMultipart('alternative')
    mime_message['From'] = email_from
    mime_message['To'] = email_to
    mime_message['Subject'] = "Top 10 Mother's Day Gifts She's Going to Love!" 
 
    # Create the body of the message (a plain-text and an HTML version).
    text = "Top 10 Mother's Day Gifts She's Going to Love!\n\n"
    text += "ThinMint's short-list of great gift ideas for Mother's Day is optimized for HTML.\n"
    text += "Please visit us online at http://www.thinmint.io/mothers-day-2015\n\n"
    text += "This email is an advertisement or solicitation. If you no longer wish to receive these emails,\n"
    text += "please visit http://www.thinmint.io/optout\n\n"
    text += "Copyright 2015 Stone's Wager, LLC dba ThinMint.io\n"
    text += "7522 Campbell Rd #113-137, Dallas, TX 75248"
 
    custom_html = customize_html(user)
    # Record the MIME types of both parts - text/plain and text/html.
    mime_part1 = MIMEText(text, 'plain')
    mime_part2 = MIMEText(custom_html, 'html')
    #mime_part2 = MIMEText(html, 'html')
 
    # Attach parts into message container. According to RFC 2046, the last part of a multipart message, in this case the HTML message, is best and preferred.
    mime_message.attach(mime_part1)
    mime_message.attach(mime_part2)
 
    #print(custom_html)
    return mime_message
 
# Get the list of all users we want to send to
users = get_user_list()
#print(users)

 
# Establish SMTP connection
server = smtplib.SMTP(host = smtp_server, port = smtp_port, timeout = 10)
server.set_debuglevel(0)
server.starttls()
server.ehlo()
server.login(smtp_username, smtp_password)
 
for user in users:
    print(user)
    userid = user[0]
    email_to = user[1]
 
    t1 = time.time()
    # Personalize the individual email
    mime_message = generate_mime_message(user)
 
    # Send the email
    smtp_response = server.sendmail(email_from, email_to, mime_message.as_string())
    # Update the user's communication status
    log_email_send(user)
    t2 = time.time()
    elapsed_time = t2 - t1
 
    print("Gen and Send: " + str(elapsed_time))
    if (elapsed_time < ses_send_rate):
        time.sleep(ses_send_rate - elapsed_time)
