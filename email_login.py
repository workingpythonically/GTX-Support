#!/usr/bin/env python3

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
# from os.path import basename

def login():
    server = smtplib.SMTP("smtp.office365.com", "587")
    server.starttls()
    server.login(username, password)
    print("Login Successful!")

def log_out():
    server.quit()
    print("Quiting the Server...")

def send_mail(accounts, Copy, Toaddress, msg):
    try:
        server.sendmail(Copy, Toaddress, msg)
        print(f"{var.upper()} Email sent {accounts}")
