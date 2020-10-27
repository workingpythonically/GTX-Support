#!/usr/bin/env python3

def SignatureType(signature):
    if signature == "GPS SmartSole":
        Copy = "smartsole@gtxcorp.com"
        team = "SmartSole Team"
        webpage = "https://gpssmartsole.com/"
        return Copy, team, webpage

    elif signature == "Invisabelt":
        Copy = "support@gtxcorp.com"
        team = "GTX Support Team"
        webpage = "https://gtxcorp.com/"
        return Copy, team, webpage

    elif signature == "Track My WorkForce":
        Copy = "tmwf@gtxcorp.com"
        team = "TMWF Team"
        webpage = "https://gtxcorp.com/tmwf-app/"
        return Copy, team, webpage

    else:
        Copy = "take-along@gtxcorp.com"
        team = "Take-Along Tracker Team"
        webpage = "https://gtxcorp.com/take-along-tracker-3g"

def SubjectType(var, Etype, notice = None):
    if var == "welcome":
        if accounts[i][0] == "GPS SmartSole":
            return f"{accounts[i][1]} - Your GPS SmartSole has been shipped!"
        elif accounts[i][0] == "Take-Along":
            return f"{accounts[i][1]} - Your Take-Along Tracker has been shipped!"
        else:
            return f"{accounts[i][1]} - Your GPS Invisabelt has been shipped!"

    if var == "refund processing"
        if accounts[i][0] == "GPS SmartSole":
            return f"{accounts[i][1]} - {accounts[i][0]} Refund Processing!"
        elif accounts[i][0] == "Invisabelt":
            return f"{accounts[i][1]} - {accounts[i][0]} Refund Processing!"
        else:
            return f"{accounts[i][1]} - {accounts[i][0]} Refund Processing!"

    if var == "declined":
        if notice == "1st" or notice == "2nd":
        return f"{OrderIDs[i]} - Payment Declined {notice} notice - {Etype} Subscription"
    else:
        return f"{OrderIDs[i]} - Service Suspended - {Etype} Subscription"

def attach_pdf():
    if accounts[i][0] == "GPS SmartSole":
        filename = "SmartSole Manual.pdf"
    elif accounts[i][0] == "Take-Along":
        filename = "Take-Along Manual.pdf"
    else:
        filename = "Invisabelt Manual.pdf"

    with open(filename, "rb") as f:
        file = MIMEApplication(f.read(), _subtype="pdf")
    file.add_header('Content-Disposition','attachment',filename=filename)
    msg.attach(file)
    f.close()
