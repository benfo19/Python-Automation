import pandas as pd
from datetime import datetime

# Set today
today = datetime.today().strftime('%m-%d-%Y')
today2 = datetime.today().strftime('%m/%d/%Y')

# Read in Backlog file, today's NAV and ORA Aging files
backlog = pd.read_excel("path/to/backlog_file.xlsx", sheet_name='Backlog Data')
nav = pd.read_excel(f"path/to/nav_aging_files/{today}.xlsx")
ora = pd.read_excel(f"path/to/ora_aging_files/{today} - Oracle Aging Report.xlsx",
                    sheet_name='Detail', skiprows=[0])

# Need to add total balance cols before filters
nav['Total Account Balance'] = nav['BalanceDue'].groupby(nav['CRMLOCID']).transform('sum')
ora['Total Account Balance'] = ora['Balance Due'].groupby(ora['Customer Number']).transform('sum')

# Add necessary filters to aging files
nav = nav.loc[(nav['Customer_Posting_Group'] != 'CUSTOMERPOSTINGGROUP') &
              (nav['AccountingRegion'] != '99') &
              (nav['AccountingRegion'] != 99) &
              (nav['BalanceDue'] < 0)]
ora = ora.loc[ora['Balance Due'] < 0]

# Compile accounts that should be closed, combine into one DF
backlog_nav = backlog[backlog['ERP'] == 'NAV']
nav_closed = backlog_nav[~backlog_nav['Entry Number'].isin(nav['EntryNo'])]

backlog_ora = backlog[backlog['ERP'] == 'ORA']
ora_closed = backlog_ora[~backlog_ora['Entry Number'].isin(ora['Unique Identifier'])]

closed_prev_day = pd.concat([nav_closed, ora_closed])

# Drop closed prev day rows from backlog data
backlog = backlog[~backlog['Entry Number'].isin(closed_prev_day['Entry Number'])]

# Now reverse, any entry numbers in NAV & ORA files not in backlog need to be added
nav_new = nav[~nav['EntryNo'].isin(backlog['Entry Number'])]
ora_new = ora[~ora['Unique Identifier'].isin(backlog['Entry Number'])]

# Rename columns to match backlog file
nav_dict = {'CompanyName': 'Business Unit',
            'CRMLOCID': 'Account Number',
            'Property_Name': 'Customer Name',
            'Document_Type': 'Document Type',
            'Document_No_': 'Document Number',
            'EntryNo': 'Entry Number',
            'External_Document_No_': 'External Ref',
            'Document_Date': 'Transaction Date',
            'BalanceDue': 'Open Balance'}
ora_dict = {'Customer Number': 'Account Number',
            'Customer Account Name': 'Customer Name',
            'Transaction Type': 'Document Type',
            'Transaction Number': 'Document Number',
            'Unique Identifier': 'Entry Number',
            'Balance Due': 'Open Balance'}

nav_new.rename(columns=nav_dict, inplace=True)
ora_new.rename(columns=ora_dict, inplace=True)

# Add ERP cols for both, currency for NAV, External Ref for ORA
nav_new['ERP'] = 'NAV'
ora_new['ERP'] = 'ORA'

nav_new['Currency'] = 'USD'
ora_new['External Ref'] = ''

# Keep only backlog columns, drop duplicate 'Customer Name' that shows up in ORA file for some reason
cols = ['ERP', 'Business Unit', 'Account Number', 'Customer Name', 'Document Type', 'Document Number', 'Entry Number',
        'Description', 'External Ref', 'Transaction Date', 'Open Balance', 'Currency', 'Total Account Balance']

nav_new = nav_new[cols]
ora_new = ora_new[cols]

ora_new = ora_new.loc[:, ~ora_new.columns.duplicated()]

# Concat ora_new and nav_new, add remaining columns
backlog_new = pd.concat([nav_new, ora_new])

backlog_new['Cash Specialist'] = ''
backlog_new['Reason Code'] = ''
backlog_new['Backlog Add Date'] = today2
backlog_new['Last Touch Date'] = ''
backlog_new['Status'] = ''
backlog_new['Note'] = ''
backlog_new['Mgmt Note'] = ''
backlog_new['Refund Status'] = ''

# Concat with remaining backlog items, Run total account balance column again to cover items that were already in the backlog
backlog = pd.concat([backlog, backlog_new])

# Save new files
path = 'path/to/save/files'

backlog.to_excel(f"{path}/Backlog/{today} Backlog.xlsx", sheet_name=f'Backlog Data {today}', index=False)
closed_prev_day.to_excel(f"{path}/Closed Prev Day/{today} Closed.xlsx", sheet_name=f'Closed Previous Day {today}', index=False)
