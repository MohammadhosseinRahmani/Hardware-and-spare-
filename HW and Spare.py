import exchangelib as Ex
import os.path
import pandas as pd
import datetime
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
import numpy as np


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
# Get the current date
current_date = datetime.datetime.now()


# Format the date as "17-Sep-2024"
formatted_date = current_date.strftime("%d-%b-%Y")
credentials = Ex.Credentials(
    username= 'mohammadhossein.rah@mtnirancell.ir',  # Or me@example.com for O365
    password='@moh321@rah654'
)
a = Ex.Account(
    primary_smtp_address='mohammadhossein.rah@mtnirancell.ir', 
    credentials=credentials, 
    autodiscover=True, 
    access_type=Ex.DELEGATE
)

TDS_Folder = a.inbox/'TDS'
# Print first 100 inbox messages in reverse order
for item in TDS_Folder.all().order_by('-datetime_received')[:1]:
    for attachment in item.attachments:
        if isinstance(attachment, Ex.FileAttachment) and attachment.name.endswith('.xlsx'):
            local_path = os.path.join("C:\\Users\\Mohammad Rahmani\\Documents\\Python\\temp", attachment.name)
            with open(local_path, "wb") as f:
                f.write(attachment.content)
            print("Saved attachment to", local_path)
            TDS = pd.read_excel(local_path,'TDS')
            print("DataFrame loaded successfully")
url = 'http://10.131.210.29/OSS_Webservices/MORNING_REPORT/uploads/morning.csv'
Morning = pd.read_csv(url, encoding='1252')
TDS = TDS[['MTTR', 'Ticket ID', 'Title', 'Site ID', 'Region', 'Province', 'FFOT', 'Processor', 'Cause', '2G', '3G', '4G', 'TDD', '5G', 'Responsible', 'Vendor']]            
TDS['RC'] = TDS['Cause'].str.upper()
TDS = TDS[~TDS['Title'].str.contains('GPS Rec')]
TDS['HardwareSpareRelated'] = np.where(
    TDS['RC'].str.contains('HARDWARE') | TDS['RC'].str.contains('SPARE') | TDS['RC'].str.contains('VANDALISM'),
    "Y", "N")
TDS = TDS[TDS['HardwareSpareRelated'] == 'Y']
TDS['MTTR'] = pd.to_timedelta(TDS['MTTR'])
HardwareIssue = pd.merge(TDS, Morning, on= 'Site ID')
HardwareIssue['Site Vendor'] = HardwareIssue.apply(
    lambda row: row['GSM Vendor'] if row['2G'] or row['3G'] or row['4G'] 
    else row['TDD Vendor'] if not (row['2G'] or row['3G'] or row['4G']) and (row['TDD']) 
    else 'Huawei', axis=1
)
HardwareIssue = HardwareIssue[['MTTR', 'Ticket ID', 'Title', 'Site ID', '5G', 'TDD', '4G', '3G', '2G', 'Region_x', 'Province_x', 'FFOT', 'Processor', 'Cause', 'GSM Vendor', 'TDD Vendor', 'Site Vendor']]

HardwareIssue['Owner'] = HardwareIssue['Processor'].apply(determine_site_vendor)
HardwareIssue.rename(columns={'Region_x':'Region', 'Province_x': 'Province', 'MTTR':'Duration', 'Processor':'Team Assigned','Cause':'Root-Cause'}, inplace=True)
Hardware_Issue = HardwareIssue[['Duration', 'Ticket ID', 'Title', 'Site ID', '5G', 'TDD', '4G', '3G', '2G', 'Region', 'Province', 'FFOT', 'Team Assigned', 'Root-Cause', 'GSM Vendor', 'TDD Vendor']]
Hardware_Issue['Duration'] = Hardware_Issue['Duration'].apply(lambda x: f"{int(x.total_seconds() // 3600)}:{int((x.total_seconds() % 3600) // 60):02d}")

pivot_table = pd.pivot_table(HardwareIssue, 
                             index=['Region','Owner','Title'] ,
                             values='Duration', 
                             aggfunc='sum')
pivot_table = pivot_table.reset_index()
pivot_table['Duration'] = pivot_table['Duration'].apply(lambda x: f"{int(x.total_seconds() // 3600)}:{int((x.total_seconds() % 3600) // 60):02d}")
with pd.ExcelWriter(f"Hardware, Spare & Material Pending cases_ {formatted_date}.xlsx",engine='openpyxl') as writer:
    Hardware_Issue.to_excel(writer, 'Hardware Issue', index=False)
    pivot_table.to_excel(writer, 'Dashboard',  index=False)
# Load the workbook and select the sheet
wb = load_workbook(f"Hardware, Spare & Material Pending cases_ {formatted_date}.xlsx")
ws = wb['Hardware Issue']

# Define the table range and create a table
table = Table(displayName="Table1", ref=ws.dimensions)

# Add a table style
style = TableStyleInfo(
    name="TableStyleMedium2", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=False
)
table.tableStyleInfo = style

# Add the table to the sheet
ws.add_table(table)
wb.save(f"Hardware, Spare & Material Pending cases_ {formatted_date}.xlsx")
# Load the workbook and select the sheet
wb = load_workbook(f"Hardware, Spare & Material Pending cases_ {formatted_date}.xlsx")
ws = wb['Dashboard']

# Define the table range and create a table
table = Table(displayName="Table2", ref=ws.dimensions)

# Add a table style
style = TableStyleInfo(
    name="TableStyleMedium2", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=False
)
table.tableStyleInfo = style

# Add the table to the sheet
ws.add_table(table)


# Save the workbook
wb.save(f"Hardware, Spare & Material Pending cases_ {formatted_date}.xlsx")
