import pandas as pd
import pyodbc

# Connect to SQL Server (Remove sensitive server and database names)
cnxn_str = ("Driver={SQL Server};"
            "Server=your_server_name;"
            "Database=your_database_name;"
            "Trusted_Connection=yes;")
cnxn = pyodbc.connect(cnxn_str)

# Queries for each of the 3 mail merges (Modify BrandCode filtering if needed)
q_friendly = ('''with cte as (
    ...
    -- Query content removed for clarity
    ...
    group by 
    [Property Name]
    ,HQLocationName
    ,CRMLOCID
    ,[Global Email Address])
    -- Query content removed for clarity
    ...
''')

q_general = ('''with cte as (
    ...
    -- Query content removed for clarity
    ...
    group by 
    [Property Name]
    ,HQLocationName
    ,HQLocationID
    ,CRMLOCID
    ,[Global Email Address])
    -- Query content removed for clarity
    ...
''')

q_severe = ('''with cte as (
    ...
    -- Query content removed for clarity
    ...
    group by 
    [Property Name]
    ,HQLocationName
    ,HQLocationID
    ,CRMLOCID
    ,[Global Email Address])
    -- Query content removed for clarity
    ...
''')

# Read data from SQL server
friendly = pd.read_sql(q_friendly, cnxn)
general = pd.read_sql(q_general, cnxn)
severe = pd.read_sql(q_severe, cnxn)

# Read in AR alignment table (different server)
cnxn_str1 = ("Driver={SQL Server};"
            "Server=your_server_name;"
            "Database=your_database_name;"
            "Trusted_Connection=yes;")
cnxn1 = pyodbc.connect(cnxn_str1)

q_alignment = ('select * from dbo.Alignment')
alignment = pd.read_sql(q_alignment,cnxn1)

# Merge AR alignment to each DataFrame
friendly = pd.merge(friendly, alignment, how = 'left', on='HQLocationID')
general = pd.merge(general, alignment, how = 'left', on='HQLocationID')
severe = pd.merge(severe, alignment, how = 'left', on='HQLocationID')

# Drop HQLocationID from each
friendly.drop(columns = ['HQLocationID'], inplace=True)
general.drop(columns = ['HQLocationID'], inplace=True)
severe.drop(columns = ['HQLocationID'], inplace=True)

# Turn Email Columns into lists
friendly['Global Email Address'] = friendly['Global Email Address'].str.split(',')
general['Global Email Address'] = general['Global Email Address'].str.split(',')
severe['Global Email Address'] = severe['Global Email Address'].str.split(',')

# Split email column into new rows, seperated by a comma
friendly = friendly.explode('Global Email Address')
general = general.explode('Global Email Address')
severe = severe.explode('Global Email Address')

# Group each DF by Analyst
friendly_grouped = friendly.groupby(friendly['ARRep'])
general_grouped = general.groupby(general['ARRep'])
severe_grouped = severe.groupby(severe['ARRep'])


# Connect to Individual AR Rep folder path (Modify folder path if needed)
path = 'your_folder_path'

# Loop each grouped df into a single Excel file for each Rep
for name in pd.concat([friendly['ARRep'], general['ARRep'], severe['ARRep']]).dropna().unique():
    try:
        with pd.ExcelWriter(f"{path}/{name}/Mail Merge Data.xlsx", mode="w", engine="xlsxwriter") as writer:
            # friendly
            friendly_df = friendly_grouped.get_group(name)
            friendly_df.to_excel(writer, sheet_name='Friendly (1-30 days)', index=False)
            # general
            general_df = general_grouped.get_group(name)
            general_df.to_excel(writer, sheet_name='General (30-90 days)', index=False)
            # severe
            severe_df = severe_grouped.get_group(name)
            severe_df.to_excel(writer, sheet_name='Severe (90+ days)', index=False)
    except:
        pass
