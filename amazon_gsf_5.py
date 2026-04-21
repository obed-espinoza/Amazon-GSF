import pandas as pd
import numpy as np

folder_outputs = r'C:\Users\it-prod-bot4\Documents\Amazon GSF\02 Outputs'

# Created by the bot
consolidated_invoices = pd.read_csv(folder_outputs+'\\Consolidated Invoices.csv')
gsf_kbs_invoices_w_colocated = pd.read_csv(folder_outputs + '\\1-gsf_kbs_invoices_w_colocated.csv')
exceptions_report = pd.read_excel(folder_outputs + '\\3-exceptions_report.xlsx')

# Filter out exceptions
gsf_kbs_invoices_w_colocated = gsf_kbs_invoices_w_colocated[
    ~gsf_kbs_invoices_w_colocated['Internal ID'].isin(exceptions_report['Internal ID'])
]

#print(exceptions_report)
#print(gsf_kbs_invoices_w_colocated)

consolidated_invoices = consolidated_invoices.dropna().drop_duplicates(['Site - WO#','Consolidated INV#']).reset_index(drop=True)

gsf_kbs_invoices_w_colocated['Site - WO#'] = gsf_kbs_invoices_w_colocated['Site Code'] + ' - ' + gsf_kbs_invoices_w_colocated['Customer WO#']


gsf_kbs_invoices_w_colocated = gsf_kbs_invoices_w_colocated.merge(consolidated_invoices, how='left', on='Site - WO#')





# Function to combine values in a group
def combine_values(x):
    # Convert all values to strings and drop duplicates
    values = x.astype(str).unique()
    # If values repeat across rows, display the first one only
    if len(values) == 1:
        return values[0]  # If all values are the same, return the single value
    # If values are numeric, add
    elif pd.api.types.is_numeric_dtype(x):
        return pd.to_numeric(sum(x))
    # else concatenate
    return ' + '.join(sorted(values))  # If values are different, join with '+'

gsf_kbs_invoices_w_colocated = gsf_kbs_invoices_w_colocated.groupby(['Site - WO#']).agg(combine_values).reset_index()

gsf_kbs_invoices_w_colocated = gsf_kbs_invoices_w_colocated.replace(['nan', 'NaN'], np.nan)
#gsf_kbs_invoices_w_colocated = gsf_kbs_invoices_w_colocated[gsf_kbs_invoices_w_colocated['Site Code'].isin(['SAX6', 'XSF2'])]

# Remove rows with missing or multiple dates
gsf_kbs_invoices_w_colocated['Contract Start'] = pd.to_datetime(gsf_kbs_invoices_w_colocated['Contract Start'], errors='coerce')
add_to_exceptions = gsf_kbs_invoices_w_colocated[gsf_kbs_invoices_w_colocated['Contract Start'].isna()]

print(gsf_kbs_invoices_w_colocated['Description'])
filtered_rows = gsf_kbs_invoices_w_colocated[gsf_kbs_invoices_w_colocated["Description"].str.contains("Parking Lot", case=False, na=False)]
add_to_exceptions = pd.concat([add_to_exceptions, filtered_rows], ignore_index=True)
filtered_rows = gsf_kbs_invoices_w_colocated[gsf_kbs_invoices_w_colocated["Description"].str.contains("quarterly sweeping", case=False, na=False)]
add_to_exceptions = pd.concat([add_to_exceptions, filtered_rows], ignore_index=True)


gsf_kbs_invoices_w_colocated = gsf_kbs_invoices_w_colocated[~gsf_kbs_invoices_w_colocated['Contract Start'].isna()]
exceptions_report = pd.concat([exceptions_report, add_to_exceptions])
exceptions_report.to_excel(folder_outputs+'\\3-exceptions_report.xlsx', index=False)



pc_combined_output = pd.read_excel(folder_outputs+'\\2-PO Lines combined_output.xlsx')

pc_template = pd.DataFrame(index=range(len(gsf_kbs_invoices_w_colocated)))

#pc_template['Site - WO#'] = gsf_kbs_invoices_w_colocated[]

gsf_kbs_invoices_w_colocated.to_excel(folder_outputs+'\\5-gsf_kbs_invoices_w_colocated_test.xlsx', index=False)
gsf_kbs_invoices_w_colocated = pd.read_excel(folder_outputs+'\\5-gsf_kbs_invoices_w_colocated_test.xlsx')

# Filter for non-empty descriptions that include either of the specified values
gsf_kbs_invoices_w_colocated = gsf_kbs_invoices_w_colocated[
    (gsf_kbs_invoices_w_colocated['Description'].notna()) &  # Not empty (not NaN)
    (gsf_kbs_invoices_w_colocated['Description'] != '') &     # Not empty string
    (gsf_kbs_invoices_w_colocated['Description'].str.contains(
        'Account/Regional Manager|Cleaning Shift Lead|General Cleaner', 
        case=True,  # Set to False if you want case-insensitive matching
        na=False
    ))
]

pc_template['Sequence Number'] = pc_template.index+1
pc_template['Invoice Number'] = gsf_kbs_invoices_w_colocated['Consolidated INV#'].fillna(gsf_kbs_invoices_w_colocated['Document Number'])
pc_template['Invoice Description'] = gsf_kbs_invoices_w_colocated['Location Number'] + ' ' + gsf_kbs_invoices_w_colocated['Description'] + ' ' + pd.to_datetime(gsf_kbs_invoices_w_colocated['Contract Start']).dt.strftime('%m/%d') + '-' + pd.to_datetime(gsf_kbs_invoices_w_colocated['Contract End']).dt.strftime('%m/%d')
pc_template['Invoice Date'] = pd.to_datetime(gsf_kbs_invoices_w_colocated['Date']).dt.strftime('%m/%d/%Y')

pc_template['Service Period Start Date'] = pd.to_datetime(gsf_kbs_invoices_w_colocated['Contract Start']).dt.strftime('%m/%d/%Y')
pc_template['Service Period End Date'] = pd.to_datetime(gsf_kbs_invoices_w_colocated['Contract End']).dt.strftime('%m/%d/%Y')
pc_template['Date of Supply'] = np.nan
pc_template['Invoice Currency'] = 'USD'
pc_template['Invoice Amount'] = gsf_kbs_invoices_w_colocated['Total'].astype(float)
pc_template['Bill to Entity Name'] = 'Amazon.com Services Inc'
pc_template['Bill To Address Country/Region Code'] = 'US'
pc_template['Ship To Address Country/Region Code'] = 'US'
pc_template['Payee Address Country/Region Code'] = 'US'
pc_template['Payee Entity Name'] = 'KellerMeyer Bulding Services'
pc_template['Invoice Type'] = 'Standard'
pc_template['Reference Invoice Number'] = np.nan
pc_template['Line Type'] = 'ITEM'
pc_template['Line Category'] = 'item_code'
pc_template['PO Number'] = gsf_kbs_invoices_w_colocated['Customer WO#']
pc_template['Line Description'] = pc_template['Invoice Description']

# Temp columns to merge and bring in Line # from PC data
pc_template['Site - WO#'] = gsf_kbs_invoices_w_colocated['Site Code'] + ' - ' + gsf_kbs_invoices_w_colocated['Customer WO#']
pc_template['Site - WO#2'] = gsf_kbs_invoices_w_colocated['replaced_colocated_site'] + ' - ' + gsf_kbs_invoices_w_colocated['Customer WO#']
pc_combined_output['Site - WO#'] = pc_combined_output['Site Name'] + ' - ' + pc_combined_output['Customer WO#']

pc_template = pc_template.merge(pc_combined_output, how='left', on='Site - WO#')
pc_template = pc_template.merge(pc_combined_output, how='left', left_on = 'Site - WO#2', right_on='Site - WO#')

pc_template['PO Line Number_x'] = pc_template['PO Line #_x']
pc_template['PO Line Number_y'] = pc_template['PO Line #_y']
pc_template['PO Line Number'] = pc_template['PO Line Number_x'].fillna(pc_template['PO Line Number_y'])
pc_template = pc_template.drop(['PO Line Number_y','PO Line Number_x','Site - WO#_x',
       'Site - WO#2', 'PO Line #_x', 'Description_x', 'Invoice Qty (?)_x',
       'Unit (?)_x', 'Price/Unit_x', 'Item Amount_x', 'Do not apply tax (?)_x',
       'Action_x', 'Customer WO#_x', 'Available Amount_x', 'Site Name_x',
       'PO Line #_y', 'Description_y', 'Invoice Qty (?)_y', 'Unit (?)_y',
       'Price/Unit_y', 'Item Amount_y', 'Do not apply tax (?)_y', 'Action_y',
       'Customer WO#_y', 'Available Amount_y', 'Site Name_y', 'Site - WO#_y'], axis=1, errors='ignore')

pc_template['Vendor Part Number'] = np.nan
pc_template['Invoiced Quantity'] = 1
pc_template['Unit of Measure'] = 'EACH'
pc_template['Unit Price'] = gsf_kbs_invoices_w_colocated['Subtotal'].astype(float)
pc_template['Line Net Amount'] = pd.to_numeric(gsf_kbs_invoices_w_colocated['Subtotal'])#.astype(float)
pc_template['Line Tax Percentage'] = round(100*(gsf_kbs_invoices_w_colocated['Tax Amount'].astype(float) / gsf_kbs_invoices_w_colocated['Subtotal'].astype(float)),2)#.replace(0, np.nan, inplace=False)
pc_template['Line Tax Amount'] = round(pd.to_numeric(gsf_kbs_invoices_w_colocated['Tax Amount']),2)#.astype(float)#.replace(0, np.nan, inplace=False)
pc_template['Is Tax Applicable'] = np.where(pc_template['Line Tax Amount'] == 0, 'NO', 'YES')
pc_template['First Attachment Name'] = np.nan
pc_template['First Attachment Description'] = np.nan
pc_template['Second Attachment Name'] = np.nan
pc_template['Second Attachment Description'] = np.nan
pc_template['Third Attachment Name'] = np.nan
pc_template['Third Attachment Description'] = np.nan

# Reorder table
pc_template = pc_template.sort_values(['Invoice Number', 'PO Number'], ascending=[True, False])

# Fix sequence number
pc_template['Sequence Number'] = (pc_template['Invoice Number'] + '_' + pc_template['PO Number'] != pc_template['Invoice Number'].shift() + '_' + pc_template['PO Number']).cumsum()

pc_template = pc_template[~pc_template['Invoice Number'].isna()]
pc_template = pc_template.replace([0, 'NaN'], np.nan)

pc_template.to_excel( folder_outputs+'\\5-pc_template.xlsx', index=False)

print("Number of columns:", len(pc_template.columns))