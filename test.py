#!/usr/bin/env python3

from email_func import SignatureType, SubjectType

# test if I use functions in a way.
accounts = [["GPS SmartSole", "702365", "Marzian", "029348103250", "mqwet@gmail.com", "werqwte41234", "gb92348"], ["Take-Along", "875046", "Zian", "098a7sdf6803250", "asdf@gail.com", "wasdfasgt4", "ga8648"]]
var = ["welcome", "refund processing"]

def Email_Loop():
    print(f"Length of Accounts: \t\t {len(accounts)}")
    for i in range(len(accounts)):
        Copy, team, webpage = SignatureType(accounts[i][0])
        # msg = MIMEMultipart()
        print(f"Copy = {Copy},\nTeam = {team},\nWebpage = {webpage}\n")
        # msg["From"] = Copy
        # msg["To"] = accounts[i][3]
        # Toaddress = [accounts[i][3], Copy]
        # msg["Subject"] = SubjectType()
        Sub = SubjectType(var[i], accounts[i][0])
        print(f"Subject: \n{Sub}\n\n")
        # Body = BodyText()
        # msg.attach(MIMEText(Body, "html"))
        # msg = msg.as_string()

    #     try:
    #         server.sendmail(username, Toaddress, msg)
    #         print(f"Welcome Email sent to {accounts[i][3]}!")
    #         print(i+1, " of ", len(accounts), "\n")

    #     except:
    #         print("Error Occured!!")
    #         print("Quiting server...")

    # server.quit()

if __name__ == '__main__':
    Email_Loop()
