#!/usr/bin/env python3
import openpyxl as xl


# might have to move the excel extraction here.
accounts = [["GPS SmartSole", "702365", "Marzian", "029348103250", "mqwet@gmail.com", "werqwte41234", "gb92348"], ["Take-Along", "875046", "Zian", "098a7sdf6803250", "asdf@gail.com", "wasdfasgt4", "ga8648"]]
var = ["welcome", "refund processing"]

def Xcel(num):
    accounts = []

    #read only, a little faster to work with.
    wb = xl.load_workbook(r'C:\Users\courier\Desktop\Automate\GTXCorp.xlsx', read_only = True)
    print("\nLoading SmartSole Workbook...")
    sheet = wb["Welcome Email"]
    num = int(input("\n\nHow many Accounts are we sending out?\n"))

    # Parsing the data into a list in a list
    for row_cells in sheet.iter_rows(min_col = 2, max_col = 9, min_row = 3, max_row = num+2):
        accounts.append([cell.value for cell in row_cells])

    # Close Workbook after reading.
    wb.close()
    print("Accessing Workbook Complete!\n\nLogging in to SmartSole Email...\n")

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

def SubjectType(var, Email_Type, OrderID, notice = None):
    if var == "welcome":
        if Email_Type == "GPS SmartSole":
            return f"{OrderID} - Your GPS SmartSole has been shipped!"
        elif Email_Type == "Take-Along":
            return f"{OrderID} - Your Take-Along Tracker has been shipped!"
        else:
            return f"{OrderID} - Your GPS Invisabelt has been shipped!"

    if var == "refund processing":
        if E == "GPS SmartSole":
            return f"{OrderID} - {Email_Type} Refund Processing!"
        elif  Email_Type == "Invisabelt":
            return f"{OrderID} - {Email_Type} Refund Processing!"
        else:
            return f"{OrderID} - {Email_Type} Refund Processing!"

    if var == "declined":
        if notice == "1st" or notice == "2nd":
            return f"{OrderID} - Payment Declined {notice} notice - {Email_type} Subscription"
    else:
        return f"{OrderID} - Service Suspended - {Email_type} Subscription"




