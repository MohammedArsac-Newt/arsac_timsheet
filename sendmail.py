import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
from datetime import datetime, timedelta

# Calculate current week's Monday and Friday dates
today = datetime.now()
monday = today - timedelta(days=today.weekday())  # Go back to Monday
friday = monday + timedelta(days=4)  # Go forward to Friday

# Format dates as required (DD-MMM-YYYY format)
from_date = monday.strftime("%d-%b-%Y").upper()  # Example: "04-NOV-2024"
to_date = friday.strftime("%d-%b-%Y").upper()    # Example: "08-NOV-2024"

# Create email.txt with automated dates
email_content = f"""From={from_date}
To={to_date}
manager=Kiran
toMailID=kirans@newtglobalcorp.com
ccMailID=timesheet@newtglobalcorp.com
"""

with open('email.txt', 'w') as file:
    file.write(email_content)

# Read variables from email.txt
variables = {}
with open('email.txt', 'r') as file:
    for line in file:
        name, value = line.strip().split('=')
        variables[name.strip()] = value.strip()

# Extract variables
from_date = variables.get('From')
to_date = variables.get('To')
manager_name = variables.get('manager')
to_mail_id = variables.get('toMailID')
cc_mail_id = variables.get('ccMailID')

# Email credentials
outlook_username = 'arsacm@newtglobalcorp.com'  # Replace with your Outlook email
outlook_password = 'Arsac@123'    # Replace with your Outlook password

# Email subject with updated date format
subject = f"Approval For Timesheet for the period from Mon {from_date} to Fri {to_date}"

# Read signature from an HTML file
with open('signature.html', 'r') as sig_file:
    signature = sig_file.read()

# Email body with signature
body = f"""
<p>Hi {manager_name},</p>
<p>I have attached the timesheet for the period from Mon {from_date} to Fri {to_date}.<br>
<span><a style="margin:0px; padding-right:1px; padding-left:1px; background-color:rgb(244,244,244); border-radius:2px; padding:0px 1px; border-radius:2px; background-color:rgb(244,244,244)" data-loopstyle="linkonly" data-ogsc="" class="x_OWAAutoLink x_eScj0 x_none" id="OLK_Beautified_307824fe-c310-7aea-b12c-33a60403ebca" data-auth="NotApplicable" rel="noopener noreferrer" target="_blank" href="https://newtglobalcorp-my.sharepoint.com/:x:/r/personal/arsacm_newtglobalcorp_com/Documents/Mohammed%20Arsac-Timesheet%20-%202025.xlsx?d=w684d072c9f7645adb9c572a06d803d2e&csf=1&web=1&e=ygo8DX" title="Original URL: https://newtglobalcorp-my.sharepoint.com/:x:/r/personal/arsacm_newtglobalcorp_com/Documents/Mohammed%20Arsac-Timesheet%20-%202025.xlsx?d=w684d072c9f7645adb9c572a06d803d2e&csf=1&web=1&e=ygo8DX. Click or tap if you trust this link." data-linkindex="0">Mohammed Arsac  -Timesheet Format - 2025.xlsx<img style="width:16px; height:16px; vertical-align:middle; padding:1px 2px 2px 0px" role="presentation" alt="" class="x_suRDx" src="https://res.public.onecdn.static.microsoft/assets/mail/file-icon/png/xlsx_16x16.png" data-imagetype="External"></a></span></p>

<p>Kindly review and approve.</p>

<!-- Add spacing before the signature -->
<p style="margin-top: 65px;"></p>

{signature}
"""

# Compose email
msg = MIMEMultipart()
msg['From'] = outlook_username
msg['To'] = to_mail_id
msg['Cc'] = cc_mail_id
msg['Subject'] = subject
msg.attach(MIMEText(body, 'html'))

# Send email
try:
    server = smtplib.SMTP('smtp.office365.com', 587)
    server.starttls()
    server.login(outlook_username, outlook_password)
    server.sendmail(outlook_username, [to_mail_id, cc_mail_id], msg.as_string())
    print("Email sent successfully.")
except Exception as e:
    print(f"Failed to send email: {e}")
finally:
    server.quit()
