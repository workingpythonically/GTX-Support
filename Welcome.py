#!/usr/bin/env python3

import smtplib
import openpyxl as xl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from os.path import basename

def BodyText():
	if accounts[i][0] == "GPS SmartSole":
		Body = f"""
<p>Hello {accounts[i][2]},&nbsp;</p>
<p>We are pleased to inform you that your GPS SmartSole&reg;&nbsp;order has been shipped.<br /><u>UPS&nbsp;tracking number</u>:&nbsp;<strong>&nbsp;{accounts[i][3]}</strong></p>
<p>To monitor your GPS SmartSole&reg;, go to:&nbsp;<a href="http://track1.gtxcorp.com/"><strong>track1.gtxcorp.com</strong></a>, and sign in with your account information.</p>
<p><strong><span style="color: #99cc00;">Your Account Information</span><br /></strong>Account:&nbsp;<strong>smartsole<br /></strong>User:<strong>&nbsp;</strong> <strong>{accounts[i][4]}<br /></strong>Password:<strong>&nbsp;{accounts[i][5]}<br /></strong>Device ID (for Smart Locator app):<strong>&nbsp;{accounts[i][6]}</strong></p>
<p>Simple-Click Login:&nbsp;<a href="https://track1.gtxcorp.com/gtxtrack/Track?page=map.device&amp;account=smartsole&amp;user={accounts[i][4]}&amp;password={accounts[i][5]}">https://track1.gtxcorp.com/gtxtrack/Track?page=map.device&amp;account=smartsole&amp;user={accounts[i][4]}&amp;password={accounts[i][5]}</a>&nbsp; &nbsp;</p>
<p><span style="color: #99cc00;"><strong>Helpful Videos &amp; Guides</strong></span></p>
<ul>
<li><a href="https://youtu.be/8vd--u9_pjE"><strong>Quick Start Tutorial Video</strong></a></li>
<li><a href="http://gpssmartsole.com/gpssmartsole/port-free-charging/"><strong>Charging Video</strong></a></li>
<li>User&rsquo;s Manual is attached (PDF file) and also included in your shipment.</li>
</ul>
<p><span style="color: #99cc00;"><strong>Please send us a brief reply confirming you have received your account information!</strong></span></p>
<p>You can sign in to your SmartSole account from any web browser by visiting the&nbsp;<a href="https://track1.gtxcorp.com/">monitoring portal link</a> above, or with our Smart Locator app from your&nbsp;<a href="https://apps.apple.com/us/app/gtx-corp-smart-locator/id767726228">Apple</a>&nbsp;or&nbsp;<a href="https://play.google.com/store/apps/details?id=com.gtxcorp.gtxcorpsmartlocator">Android</a>&nbsp;smart devices. For the Smart Locator app, refer to the Device ID above when adding your SmartSole for the first time.</p>
<p><span style="color: #99cc00;"><strong>Download Smart Locator App</strong></span></p>
<ul>
<li><a href="https://itunes.apple.com/us/app/gtx-corp-smart-locator/id767726228?mt=8"><strong>Apple iPhone/iPad</strong></a></li>
<li><strong><a href="https://play.google.com/store/apps/details?id=com.gtxcorp.gtxcorpsmartlocator">Android smartphone or tablet</a></strong></li>
</ul>
<p><strong><span style="color: #99cc00;">Coverage Notice</span><br /></strong>PLEASE NOTE: Your&nbsp;GPS SmartSole&reg;&nbsp;operates like a cell phone and&nbsp;<u>requires sufficient 2G coverage</u>&nbsp;to function properly. While 2G is still available today, it is planned for eventual sunset by some operators (T-Mobile in US and Rogers in Canada). Therefore, you must test your GPS SmartSole for satisfactory cellular connectivity BEFORE trimming or testing it for comfort. If you are not totally satisfied with the network performance,&nbsp;please&nbsp;notify us within 14 days of receiving the product. Any returns after 14 days will be subject to a 20% restocking fee. Trimmed or worn product is not eligible for return or refund.</p>
<p>For more information about your purchase, including the full warranty and return policies, please refer to the&nbsp;<a href="https://gpssmartsole.com/gpssmartsole/terms-conditions-of-purchase/">Terms and Conditions</a>:<br /><a href="https://gpssmartsole.com/gpssmartsole/terms-conditions-of-purchase/"><strong>https://gpssmartsole.com/gpssmartsole/terms-conditions-of-purchase/</strong></a></p>
<p>Please let us know if you would like any assistance setting up your GPS SmartSole&reg;, as we are happy to schedule an appointment. The quickest way to reach us for support or billing inquiries is to email&nbsp;<strong><a href="mailto:smartsole@gtxcorp.com">smartsole@gtxcorp.com</a>.</strong></p>
<p>Sincerely,</p>
<h3><strong>{team}</strong></h3>
<p><span style="color: #ff0000;"><strong>GTX Corp</strong><br /></span>117 West 9th Street, Suite 1214<strong><br /></strong>Los Angeles, California 90015&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;<br /><a href="mailto:{Copy}">{Copy}<br /></a><a href="{webpage}">{webpage}</a><a href="https://gpssmartsole.com/"><br /></a>www.GTXCorp.com&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p>
<p>(OTCQB:&nbsp;GTXO)<br /><span style="color: #0000ff;"><strong><a href="https://www.youtube.com/watch?v=R3Ssy1mnDcA">https://www.youtube.com/watch?v=R3Ssy1mnDcA</a><a href="https://www.youtube.com/watch?v=R3Ssy1mnDcA"><br /></a></strong></span>&nbsp;This message contains confidential information and is intended only for {accounts[i][2]}. If you are not the named addressee you should not disseminate, distribute or copy this e-mail. Please notify the sender immediately if you have received this e-mail by mistake and delete this e-mail from your system. Finally, the recipient should check this email and any attachments for the presence of viruses. The company accepts no liability for any damage caused by any virus transmitted by this email.</p>
"""
		return Body
	elif accounts[i][0] == "Take-Along":
		Body = f"""
<p>Hello {accounts[i][2]},&nbsp;</p>
<p>We are pleased to inform you that your Take-Along Tracker order has been shipped: <strong>{accounts[i][3]}</strong></p>
<p>To monitor your Take-Along Tracker&reg;, go to:&nbsp;<a href="http://track1.gtxcorp.com/"><strong>track1.gtxcorp.com</strong></a>, and sign in with your account information.</p>
<p><strong><span style="color: #99cc00;">Your Account Information</span><br /></strong>Account: Take-Along<strong><br /></strong>User:<strong>&nbsp;</strong>&nbsp;<strong>{accounts[i][4]}<br /></strong>Password:<strong>&nbsp;{accounts[i][5]}<br /></strong>Device ID (for Smart Locator app):<strong>&nbsp;{accounts[i][6]}</strong></p>
<p>Simple-Click Login:&nbsp;<a href="https://track1.gtxcorp.com/gtxtrack/Track?page=map.device&amp;account=take-along&amp;user={accounts[i][4]}&amp;password={accounts[i][5]}">https://track1.gtxcorp.com/gtxtrack/Track?page=map.device&amp;account=take-along&amp;user={accounts[i][4]}&amp;password={accounts[i][5]}</a>&nbsp; &nbsp;</p>
<p><span style="color: #99cc00;"><strong>Please send us a brief reply confirming you have received your account information!</strong></span></p>
<p>You can sign in to your&nbsp;Take-Along account from any web browser by visiting the&nbsp;<a href="https://track1.gtxcorp.com/">monitoring portal link</a>&nbsp;above, or with our Smart Locator app from your&nbsp;<a href="https://apps.apple.com/us/app/gtx-corp-smart-locator/id767726228">Apple</a>&nbsp;or&nbsp;<a href="https://play.google.com/store/apps/details?id=com.gtxcorp.gtxcorpsmartlocator">Android</a>&nbsp;smart devices. For the Smart Locator app, refer to the Device ID above when adding your&nbsp;Take-Along Tracker for the first time.</p>
<p><span style="color: #99cc00;"><strong>Download Smart Locator App</strong></span></p>
<ul>
<li><a href="https://itunes.apple.com/us/app/gtx-corp-smart-locator/id767726228?mt=8"><strong>Apple iPhone/iPad</strong></a></li>
<li><strong><a href="https://play.google.com/store/apps/details?id=com.gtxcorp.gtxcorpsmartlocator">Android smartphone or tablet</a></strong>&nbsp;</li>
</ul>
<p>Please see attached for the User&rsquo;s Guide. Please note the pictured unit has a charging cradle, and the product no longer supports a charging cradle &ndash; it now charges directly via the included charging cable, for a better connection.</p>
<p>When using the portal, please note that &ldquo;Locate&rdquo; and &ldquo;Battery&rdquo; functions can incur additional costs. Refer to the attached &ldquo;Value Added Features&rdquo; document for details on the monthly quota and additional costs.</p>
<p>Please let us know if you have any questions regarding your Take-Along Tracker.</p>
<p>Sincerely,</p>
<h3><strong>{team}</strong></h3>
<p><span style="color: #ff0000;"><strong>GTX Corp</strong><br /></span>117 West 9th Street, Suite 1214<strong><br /></strong>Los Angeles, California 90015&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;<br /><a href="mailto:{Copy}">{Copy}<br /></a><a href="{webpage}">{webpage}</a><a href="https://gpssmartsole.com/"><br /></a>www.GTXCorp.com&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p>
<p>(OTCQB:&nbsp;GTXO)<br /><span style="color: #0000ff;"><strong><a href="https://www.youtube.com/watch?v=R3Ssy1mnDcA">https://www.youtube.com/watch?v=R3Ssy1mnDcA</a><a href="https://www.youtube.com/watch?v=R3Ssy1mnDcA"><br /></a></strong></span>&nbsp;This message contains confidential information and is intended only for {accounts[i][2]}. If you are not the named addressee you should not disseminate, distribute or copy this e-mail. Please notify the sender immediately if you have received this e-mail by mistake and delete this e-mail from your system. Finally, the recipient should check this email and any attachments for the presence of viruses. The company accepts no liability for any damage caused by any virus transmitted by this email.</p>
"""
		return Body
	else:
		Body = f"""
<p>Hello {accounts[i][2]},&nbsp;</p>
<p>We are pleased to inform you that your GPS Invisabelt&reg;&nbsp;order has been shipped.<br /><strong><span style="color: #3366ff;"><u>Tracking number</u>:&nbsp;</span>&nbsp;{accounts[i][3]}</strong></p>
<p>To monitor your GPS Invisabelt&reg;, go to:&nbsp;<a href="http://track1.gtxcorp.com/"><strong>track1.gtxcorp.com</strong></a>, and sign in with your account information.</p>
<p><strong><span style="color: #3366ff;">Your Account Information</span><br /></strong>Account:&nbsp;<strong>Invisabelt<br /></strong>User:<strong>&nbsp;</strong> <strong>{accounts[i][4]}<br /></strong>Password:<strong>&nbsp;{accounts[i][5]}<br /></strong>Device ID (for Smart Locator app):<strong>&nbsp;{accounts[i][6]}</strong></p>
<p>Simple-Click Login:&nbsp;<a href="https://track1.gtxcorp.com/gtxtrack/Track?page=map.device&amp;account=invisabelt&amp;user=user={accounts[i][4]}&amp;password={accounts[i][5]}">https://track1.gtxcorp.com/gtxtrack/Track?page=map.device&amp;account=invisabelt&amp;user={accounts[i][4]}&amp;password={accounts[i][5]}</a>&nbsp; &nbsp;</p>
<p><span style="color: #3366ff;"><strong>Please send us a brief reply confirming you have received your account information!</strong></span></p>
<p>You can sign in to your&nbsp;Invisabelt account from any web browser by visiting the&nbsp;<a href="https://track1.gtxcorp.com/">monitoring portal link</a> above, or with our Smart Locator app from your&nbsp;<a href="https://apps.apple.com/us/app/gtx-corp-smart-locator/id767726228">Apple</a>&nbsp;or&nbsp;<a href="https://play.google.com/store/apps/details?id=com.gtxcorp.gtxcorpsmartlocator">Android</a>&nbsp;smart devices. For the Smart Locator app, refer to the Device ID above when adding your&nbsp;Invisabelt for the first time.</p>
<p><span style="color: #3366ff;"><strong>Download Smart Locator App</strong></span></p>
<ul>
<li><span style="color: #000000;"><a style="color: #000000;" href="https://itunes.apple.com/us/app/gtx-corp-smart-locator/id767726228?mt=8">Apple iPhone/iPad</a></span></li>
<li><span style="color: #000000;"><a style="color: #000000;" href="https://play.google.com/store/apps/details?id=com.gtxcorp.gtxcorpsmartlocator">Android smartphone or tablet</a></span></li>
</ul>
<p><strong><span style="color: #99cc00;"><span style="color: #3366ff;">Coverage Notice</span><br /></span></strong>Your&nbsp;Invisabelt&nbsp;shipment includes a&nbsp;User&rsquo;s Manual. The User&rsquo;s Manual is also attached in this email (PDF file). YOU MUST TEST YOUR GPS INVISABELT FOR CONNECTIVITY IN YOUR AREA. Follow test instructions included in the User's Manual.</p>
<p>For reference, our Terms and Conditions, which includes our warranty and return policy, is available here:&nbsp;<a href="http://www.gpsinvisabelt.com/terms-conditions-of-sale/">http://www.gpsinvisabelt.com/terms-conditions-of-sale/</a></p>
<p>Please let us know if you would like any assistance setting up your GPS Invisabelt&reg;, as we are happy to schedule an appointment. The quickest way to reach us for support or billing inquiries is to email&nbsp;<strong><a href="mailto:smartsole@gtxcorp.com">support@gtxcorp.com</a>.</strong></p>
<p>Sincerely,</p>
<h3><strong>{team}</strong></h3>
<p><span style="color: #ff0000;"><strong>GTX Corp</strong><br /></span>117 West 9th Street, Suite 1214<strong><br /></strong>Los Angeles, California 90015&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;<br /><a href="mailto:{Copy}">{Copy}<br /></a><a href="{webpage}">{webpage}</a><a href="https://gpssmartsole.com/"><br /></a>www.GTXCorp.com&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p>
<p>(OTCQB:&nbsp;GTXO)<br /><span style="color: #0000ff;"><strong><a href="https://www.youtube.com/watch?v=R3Ssy1mnDcA">https://www.youtube.com/watch?v=R3Ssy1mnDcA</a><a href="https://www.youtube.com/watch?v=R3Ssy1mnDcA"><br /></a></strong></span>&nbsp;This message contains confidential information and is intended only for {accounts[i][2]}. If you are not the named addressee you should not disseminate, distribute or copy this e-mail. Please notify the sender immediately if you have received this e-mail by mistake and delete this e-mail from your system. Finally, the recipient should check this email and any attachments for the presence of viruses. The company accepts no liability for any damage caused by any virus transmitted by this email.</p>
"""
		return Body

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
		webpage = "https://gtxcorp.com/take-along-tracker-3g/"
		return Copy, team, webpage

def SubjectType():
	if accounts[i][0] == "GPS SmartSole":
		return f"{accounts[i][1]} - Your GPS SmartSole has been shipped!"
	elif accounts[i][0] == "Take-Along":
		return f"{accounts[i][1]} - Your Take-Along Tracker has been shipped!"
	else:
		return f"{accounts[i][1]} - Your GPS Invisabelt has been shipped!"

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


username = "mtulabot@gtxcorp.com"
password = "Mrt012048554"

wb = xl.load_workbook(r'C:\Users\courier\Desktop\Automate\GTXCorp.xlsx') #, read_only = True)
print("\nLoading SmartSole Workbook...")
sheet = wb["Welcome Email"]

num = int(input("\n\nHow many Accounts are we sending out?\n"))
accounts = [] # this is where we will store the list of values from the Excel sheet

# Slightly different from PaymentNotice.py, this iterates the whole table B3:H5
# B is column 2 on excel and H is column 8 on excel.
# This code to extract the Excel sheet is faster because it only uses 1 loop
for row_cells in sheet.iter_rows(min_col = 2, max_col = 9, min_row = 3, max_row = num+2):
	accounts.append([cell.value for cell in row_cells])

print("Accessing Workbook Complete!\n\nLogging in to SmartSole Email...\n")

server = smtplib.SMTP("smtp.office365.com", "587")
server.starttls()
server.login(username, password)

print("Login Successful!")

"""
accounts[i][0] = Account ex "SmartSole", "Invisabelt", "Take-Along"
accounts[i][1] = Order ID
accounts[i][2] = Name
accounts[i][3] = Tracking number
accounts[i][4] = Email
accounts[i][5] = Serial / password
accounts[i][6] = Device ID
"""

for i in range(len(accounts)):
	Copy, team, webpage = SignatureType()
	msg = MIMEMultipart()
	msg["From"] = Copy
	msg["To"] = accounts[i][4]
	Toaddress = [accounts[i][4], Copy]
	msg["Subject"] = SubjectType()
	Body = BodyText()
	msg.attach(MIMEText(Body, "html"))

	attach_pdf()
	msg = msg.as_string()
	print("Attachement Successful!")

	try:
		server.sendmail(Copy, Toaddress, msg)
		print(f"Welcome Email sent to {accounts[i][4]}!")
		print(i+1," of ",len(accounts),"\n")

	except:
		print("Error Occured!")
		print("Quiting server..")
		break

server.quit()
print("Quiting the Server...")
print(f"All emails sent successfully!\n\t\tTotal emails sent: {len(accounts)}\n")
