#!/usr/bin/env python3

import smtplib
import openpyxl as xl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from os.path import basename

def SignatureType():
    if accounts[i][0] == "GPS SmartSole":
        Copy = "smartsole@gtxcorp.com"
        team = "SmartSole Team"
        webpage = "https://gpssmartsole.com/"
        return Copy, team, webpage

    elif accounts[i][0] == "Invisabelt":
        Copy = "support@gtxcorp.com"
        team = "GTX Support Team"
        webpage = "https://gtxcorp.com/"
        return Copy, team, webpage

    else:
        Copy = "take-along@gtxcorp.com"
        team = "Take-Along Tracker Team"
        webpage = "https://gtxcorp.com/take-along-tracker-3g"

def SubjectType():
    if accounts[i][0] == "GPS SmartSole":
        return f"{accounts[i][1]} - {accounts[i][0]} Refund Processing!"
    elif accounts[i][0] == "Invisabelt":
        return f"{accounts[i][1]} - {accounts[i][0]} Refund Processing!"
    else:
        return f"{accounts[i][1]} - {accounts[i][0]} Refund Processing!"

def BodyText():
    Body = f"""
<p>Dear {accounts[i][2]},</p>
<p>After thorough testing and inpection, we are pleased to inform you that your {accounts[i][0]} has passed all qualifications for refund.</p>
<p>Our Billing Department will provide an update to let you know when the refund has been issued. Please allow up to 1-2 billing cycles to fully process your refund. If you have any questions regarding your refund, please contact our Billing Department: <a href="mailto:billing@gtxcorp.com">billing@gtxcorp.com</a></p>
<p>We sincerely hope that you can find your way back to GTX Corp someday.&nbsp;</p>
<p>Sincerely,</p>
<h3><strong>{team}</strong></h3>
<p><span style="color: #ff0000;"><strong>GTX Corp</strong><br /></span>117 West 9th Street, Suite 1214<strong><br /></strong>Los Angeles, California 90015&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;<br /><a href="mailto:{Copy}">{Copy}<br /></a><a href="{webpage}">{webpage}</a><a href="https://gpssmartsole.com/"><br /></a>www.GTXCorp.com&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p>
<p>(OTCQB:&nbsp;GTXO)<br /><span style="color: #0000ff;"><strong><a href="https://www.youtube.com/watch?v=R3Ssy1mnDcA">https://www.youtube.com/watch?v=R3Ssy1mnDcA</a><a href="https://www.youtube.com/watch?v=R3Ssy1mnDcA"><br /></a></strong></span>&nbsp;This message contains confidential information and is intended only for {accounts[i][2]}. If you are not the named addressee you should not disseminate, distribute or copy this e-mail. Please notify the sender immediately if you have received this e-mail by mistake and delete this e-mail from your system. Finally, the recipient should check this email and any attachments for the presence of viruses. The company accepts no liability for any damage caused by any virus transmitted by this email.</p>
"""
    return Body


wb = xl.load_workbook(r'C:\Users\courier\Desktop\Automate\GTXCorp.xlsx')
print("\nLoading GTX Workbook...")
sheet = wb["Refund Processing"]
num = int(input("\n\nHow many accounts are we sending out? \n"))
accounts = []

for row_cells in sheet.iter_rows(min_col = 2, max_col = 5, min_row = 3, max_row = num+2):
    accounts.append([cell.value for cell in row_cells])

print("Accessing Workbook Complete! \n\n Logging in to GTX Email...\n")

'''
accounts[i][0] = Account
accounts[i][1] = Order ID
accounts[i][2] = Name
accounts[i][3] = Email
'''

username = "mtulabot@gtxcorp.com"
password = "Mrt012048554"

server = smtplib.SMTP("smtp.office365.com", "587")
server.starttls()
server.login(username, password)

print("Login Successful!")

for i in range(len(accounts)):
    Copy, team, webpage = SignatureType()
    msg = MIMEMultipart()
    msg["From"] = Copy
    msg["To"] = accounts[i][3]
    Toaddress = [accounts[i][3], Copy]
    msg["Subject"] = SubjectType()
    Body = BodyText()
    msg.attach(MIMEText(Body, "html"))
    msg = msg.as_string()

    try:
        server.sendmail(username, Toaddress, msg)
        print(f"Welcome Email sent to {accounts[i][3]}!")
        print(i+1, " of ", len(accounts), "\n")

    except:
        print("Error Occured!!")
        print("Quiting server...")

server.quit()
print("Quiting the server...")
print(f"All emails send successfully! \n\t\t Total emails sent: {len(accounts)}\n")
