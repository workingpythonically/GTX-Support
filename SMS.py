# GOAL 1. Send SMS to phone
# Idea: send a message using smtplib (Simple Mail Transfer Protocol) found help online

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# This is where wer are sending from
email = "mtulabot@gtxcorp.com" #/smartsole@gtxcorp.com/"
pw = "Mrt_012048554"

sms_gateway = "marziantulabot@live.com" #Marzian's number. This is where we're sending to
smtp = "smtp.office365.com" #/mail/smartsole@gtxcorp.com/" #"smtp.office365.com"
port = "587" #IDK how to find my port?

server = smtplib.SMTP(smtp,port)
server.starttls()
server.login(email,pw)

msg = MIMEMultipart()
msg["From"] = email
msg["To"] = sms_gateway

msg["Subject"] = "SMS Test"
body = "Hello!\n BITCHHHH! MAAAMOOOON!!\n Chiquitita pero Grandota!"

msg.attach(MIMEText(body, "plain"))

sms = msg.as_string()
server.SentOnBehalfOfName = "smartsole@gtxcorp.com"
server.sendmail(email, sms_gateway, sms)
server.quit()
