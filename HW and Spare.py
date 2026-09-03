# ==============================================
# MTNi Hardware & Spare Pending Cases Report Automation
# ==============================================
# Author: Mohammad Rahmani
# Purpose: Automates downloading Exchange email attachments, processes the data,
#          creates Excel reports with pivot tables, and prepares an Outlook email.
# ==============================================

import exchangelib as Ex
import os
import pandas as pd
import datetime
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
import numpy as np
import win32com.client as win32

# -----------------------------
# Function: Determine Site Vendor
# -----------------------------
# Maps the 'Processor' field to its respective team/vendor
def determine_site_vendor(processor):
    if 'MS_HD' in processor or 'RM_' in processor or 'MS_FO' in processor:
        return 'MS FM'
    elif 'MTN_NWG_Hardware' in processor:
        return 'MTN HWS'
    elif 'MS_Hardware' in processor:
        return 'MS HWS'
    elif 'OPS' in processor or 'MTNi' in processor:
        return 'OPS'
    elif 'CPG' in processor:
        return 'CPG'
    else:
        return ''

# -----------------------------
# Step 1: Setup
# -----------------------------
current_date = datetime.datetime.now()
formatted_date = current_date.strftime("%d-%b-%Y")  # e.g., "03-Oct-2025"

# Exchange email credentials
credentials = Ex.Credentials(
    username='Email',  # Change to your Exchange email
    password='Password' # Change to your password
)

# Log into Exchange account
account = Ex.Account(
    primary_smtp_address='Email', 
    credentials=credentials, 
    autodiscover=True, 
    access_type=Ex.DELEGATE
)

# -----------------------------
# Step 2: Download the latest Excel attachment from a specific folder
# -----------------------------
TDS_Folder = account.inbox / 'Folder Name'  # Folder in Exchange inbox

# Loop through the latest email(s)
for item in TDS_Folder.all().order_by('-datetime_received')[:1]:
    for attachment in item.attachments:
        if isinstance(attachment, Ex.FileAttachment) and attachment.name.endswith('.xlsx'):
            local_path = os.path.join(os.getcwd(), "Download", attachment.name)  # Save in "Download" folder beside script
            os.makedirs(os.path.dirname(local_path), exist_ok=True)
            with open(local_path, "wb") as f:
                f.write(attachment.content)
            print("Saved attachment to", local_path)
            TDS = pd.read_excel(local_path,'TDS')
            print("DataFrame loaded successfully")

# -----------------------------
# Step 3: Load network data from CSV
# -----------------------------
url = 'Network data CSV Link'  # Replace with actual link
Morning = pd.read_csv(url, encoding='1252')

# -----------------------------
# Step 4: Clean & filter TDS data
# -----------------------------
TDS = TDS[['MTTR', 'Ticket ID', 'Title', 'Site ID', 'Region', 'Province', 'FFOT', 'Processor', 'Cause', '2G', '3G', '4G', 'TDD', '5G', 'Responsible', 'Vendor']]            
TDS['RC'] = TDS['Cause'].str.upper()
TDS = TDS[~TDS['Title'].str.contains('GPS Rec')]  # Exclude GPS Recovery Tickets
TDS['HardwareSpareRelated'] = np.where(
    TDS['RC'].str.contains('HARDWARE|SPARE|VANDALISM'),
    "Y", "N"
)
TDS = TDS[TDS['HardwareSpareRelated'] == 'Y']
TDS['MTTR'] = pd.to_timedelta(TDS['MTTR'])

# Merge with network data
HardwareIssue = pd.merge(TDS, Morning, on='Site ID')

# Determine site vendor
HardwareIssue['Site Vendor'] = HardwareIssue.apply(
    lambda row: row['GSM Vendor'] if row['2G'] or row['3G'] or row['4G'] 
    else row['TDD Vendor'] if not (row['2G'] or row['3G'] or row['4G']) and (row['TDD']) 
    else 'Huawei', axis=1
)

# Select required columns
HardwareIssue = HardwareIssue[['MTTR', 'Ticket ID', 'Title', 'Site ID', '5G', 'TDD', '4G', '3G', '2G', 'Region_x', 'Province_x', 'FFOT', 'Processor', 'Cause', 'GSM Vendor', 'TDD Vendor', 'Site Vendor']]

# Map Processor to Owner
HardwareIssue['Owner'] = HardwareIssue['Processor'].apply(determine_site_vendor)

# Rename columns for final report
HardwareIssue.rename(columns={
    'Region_x':'Region', 
    'Province_x':'Province', 
    'MTTR':'Duration', 
    'Processor':'Team Assigned', 
    'Cause':'Root-Cause'
}, inplace=True)

# Format Duration as H:M
Hardware_Issue = HardwareIssue[['Duration', 'Ticket ID', 'Title', 'Site ID', '5G', 'TDD', '4G', '3G', '2G', 'Region', 'Province', 'FFOT', 'Team Assigned', 'Root-Cause', 'GSM Vendor', 'TDD Vendor']]
Hardware_Issue['Duration'] = Hardware_Issue['Duration'].apply(lambda x: f"{int(x.total_seconds() // 3600)}:{int((x.total_seconds() % 3600) // 60):02d}")

# -----------------------------
# Step 5: Create Pivot Table for Dashboard
# -----------------------------
pivot_table = pd.pivot_table(HardwareIssue, 
                             index=['Region','Owner','Title'],
                             values='Duration', 
                             aggfunc='sum').reset_index()

pivot_table['Duration'] = pivot_table['Duration'].apply(lambda x: f"{int(x.total_seconds() // 3600)}:{int((x.total_seconds() % 3600) // 60):02d}")

# -----------------------------
# Step 6: Save to Excel with tables
# -----------------------------
output_file = f"Hardware, Spare & Material Pending cases_ {formatted_date}.xlsx"
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    Hardware_Issue.to_excel(writer, 'Hardware Issue', index=False)
    pivot_table.to_excel(writer, 'Dashboard', index=False)

# Format "Hardware Issue" sheet as table
wb = load_workbook(output_file)
ws = wb['Hardware Issue']
table = Table(displayName="Table1", ref=ws.dimensions)
style = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
table.tableStyleInfo = style
ws.add_table(table)

# Format "Dashboard" sheet as table
ws = wb['Dashboard']
table = Table(displayName="Table2", ref=ws.dimensions)
table.tableStyleInfo = style
ws.add_table(table)

wb.save(output_file)

# Convert pivot table to HTML for email
Table_HTML = pivot_table.to_html(index=False)

# -----------------------------
# Step 7: Prepare Email
# -----------------------------
def Email_preparation():
    subject = f'Hardware, Spare & Material Pending cases_ {formatted_date}'
    body = f"""   
    <html>
    <body>
        <p>
        Dear RMs, <br>
            &emsp; Please support updating and reducing MS domain, especially for Huawei vendor pending cases.<br>
            For those out of MS hand, please assign them to the owner counterpart from MTNi and remove from MS basket.<br><br>
        Dear Amin/Spare Team, <br>
            &emsp; For Spare part pending cases also, would be appreciated for support.<br><br>
        </p>
        <style>
            table {{
                width: 15%;
                border: 2px solid #2E8B57;
            }}
            th {{
                background-color: #008B8B;
                color: #ffffff;
                padding:8px;
                text-align: center;
                font-size: 14px;
            }}
            td {{
                border: 1px solid #dddddd;
                padding: 9px;
                text-align: center;
                font-size: 11px;
            }}
            tr {{
                background: #F0FFFF;
            }}
        </style>
        {Table_HTML}
    </body>
    </html>
    """

    # Create Outlook email
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.HTMLBody = body
    mail.To = 'Email TO'
    mail.CC = 'Email CC'
    mail.Attachments.Add(os.path.join(os.getcwd(), "Attached", output_file))
    mail.Display()  # Display email before sending

# Send/preview email
Email_preparation()
