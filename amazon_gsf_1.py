import pandas as pd

# --- SETTINGS ---
pd.set_option("display.max_rows", None, "display.max_columns", None)

# --- LOAD DATA ---
file_path = r'C:\Users\it-prod-bot4\Documents\Amazon GSF\01 Inputs\InvoicesCreditsReadyforConsolidationAmazonResults.csv'
icrfc = pd.read_csv(file_path)

# --- CONVERT DATE COLUMNS ---
date_cols = ['Date', 'Contract Start', 'Contract End']
for col in date_cols:
    icrfc[col] = pd.to_datetime(icrfc[col])

# --- CONVERT CURRENCY TO NUMERIC ---
def clean_currency(value):
    """Removes $, commas, and handles accounting parentheses (e.g., ($10) -> -10)."""
    val_str = str(value).replace('$', '').replace(',', '')
    if '(' in val_str:
        val_str = '-' + val_str.replace('(', '').replace(')', '')
    return float(val_str)

currency_cols = ['Subtotal', 'Tax Amount', 'Total']
for col in currency_cols:
    icrfc[col] = icrfc[col].apply(clean_currency)

# --- FILTER BY DATE (LAST 28 DAYS) ---
# normalize() sets time to midnight to ensure full-day coverage
today = pd.Timestamp.today().normalize()
start_date = today - pd.DateOffset(days=28)  # Today + 27 prior days = 28 day window

icrfc = icrfc[icrfc['Date'].between(start_date, today)]

# --- FILTER BY CONTRACT MONTH (must span the full prior calendar month) ---
# Keep only invoices whose Contract Start is the first day of last month AND
# whose Contract End is the last day of last month. Anything that does not span
# the exact month being billed is removed from submission (per guide).
first_of_current_month = today.replace(day=1)
last_month_end = first_of_current_month - pd.Timedelta(days=1)
last_month_start = last_month_end.replace(day=1)

contract_mask = (
    (icrfc['Contract Start'].dt.normalize() == last_month_start) &
    (icrfc['Contract End'].dt.normalize() == last_month_end)
)
print(f"Contract month window: {last_month_start.date()} to {last_month_end.date()}")
print(f"Rows dropped by contract-month filter: {int((~contract_mask).sum())}")
icrfc = icrfc[contract_mask]


# --- OUTPUTS ---
print(f"Filtering for range: {start_date.date()} to {today.date()}")
print(f"Number of rows after filter: {len(icrfc)}")
print(icrfc.head())


# --- SCOPE TO THE AMAZON GSF (KBS) AND AMAZON IFS ENTITIES ---
# Selection is by parent entity, NOT by 'Customer Billing Group' (which may be
# empty for newly added sites). Keep Amazon - GSF (KBS) = P10133 and Amazon
# (IFS) = P13792. Everything else (e.g. System 1 Amazon P12501) is dropped.
amazon_gsf_parents = ['P10133', 'P13792']
icrfc = icrfc[icrfc['Top Level Parent'].astype(str).str.contains('|'.join(amazon_gsf_parents), na=False)]

#print(icrfc['Customer Billing Group'])

print('number of rows: '+str(len(icrfc)))

# Read excel file for colocated sites and load into dataframe
file_path = r'C:\Users\it-prod-bot4\Documents\Amazon GSF\01 Inputs\Co-located Sites Amazon GSF.xlsx'
co_located_sites = pd.read_excel(file_path)

co_located_sites.head()

# Read excel file for SO Summary and load into dataframe (for filling in customer billing groups)
#file_path = r'C:\Users\it-prod-bot4\Documents\Amazon GSF\01 Inputs\Amazon SO Summary.xlsx'
file_path = r'C:\Users\it-prod-bot4\Kellermeyer Bergensons Services, LLC\IT Production Bot - Amazon GSF SO Summary\Amazon SO Summary.xlsx'
#file_path = r'C:\Users\it-prod-bot4\OneDrive - Kellermeyer Bergensons Services, LLC\Documents\Amazon GSF SO Summary\Amazon SO Summary.xlsx'
so_summary = pd.read_excel(file_path, header=6)

so_summary.head()

"""###Replace customer billing groups with valid values"""

#from re import A
# Merge the DataFrames to bring in missing customer billing groups
merged_df = pd.merge(icrfc, so_summary, left_on='Location Number', right_on='Site Code', how='left')
#print('merged_df')
#print(merged_df['Customer Billing Group'])
# Check how many customer billing groups are empty
#merged_df = merged_df[merged_df['Customer Billing Group'].isna()]
#print(len(merged_df))

print('number of rows merged_df: '+str(len(merged_df)))

merged_df.head()

"""### Replace customer billing groups from so summary and filter for GSF KBS"""

# Replace empty 'Customer Billing Group' with 'Site Type'
merged_df['Customer Billing Group'] = merged_df['Customer Billing Group'].fillna(merged_df['Site Type'])
#print('merged_df2')
#print(merged_df['Customer Billing Group'])
# --- SELECT GSF/IFS SCOPE (SO Summary is the source of truth) ---
# Do NOT gate on 'Customer Billing Group' - it may be empty for newly added
# sites. Scope was already limited to the Amazon GSF (KBS, P10133) and Amazon
# IFS (P13792) parents above. Any site not present in the SO Summary output is
# dropped here (silently), per the guide.
before_so = len(merged_df)
merged_df = merged_df[merged_df['Site Code'].notna()].reset_index(drop=True)
print(f"Rows dropped (site not found in SO Summary): {before_so - len(merged_df)}")

gsf_kbs_invoices = merged_df.reset_index(drop=True)

gsf_kbs_invoices.head()


print('number of rows gsf_kbs_invoices: '+str(len(gsf_kbs_invoices)))

"""### Add colocated sites"""

# Merge gsf_kbs_invoices with co_located_sites
#print('e')
#print(gsf_kbs_invoices)
#print(co_located_sites.columns)
gsf_kbs_invoices_w_colocated = pd.merge(gsf_kbs_invoices, co_located_sites.add_prefix('co_located_sites_'), left_on='Location Number', right_on='co_located_sites_Site', how='left')
gsf_kbs_invoices_w_colocated[gsf_kbs_invoices_w_colocated['co_located_sites_Site'].notna()].head()
#print(gsf_kbs_invoices_w_colocated)


"""### Fill missing colocated sites with location number"""

gsf_kbs_invoices_w_colocated['replaced_colocated_site'] = gsf_kbs_invoices_w_colocated['co_located_sites_Colocated'].fillna(gsf_kbs_invoices['Location Number'])
gsf_kbs_invoices_w_colocated = gsf_kbs_invoices_w_colocated.drop(['co_located_sites_Site', 'co_located_sites_Colocated'], axis=1)
gsf_kbs_invoices_w_colocated.head()

"""### Get subtotal by site to later compare against Payee Central"""

# gsf_kbs_subtotal by site
#print(gsf_kbs_invoices_w_colocated['replaced_colocated_site'])
gsf_kbs_subtotal = gsf_kbs_invoices_w_colocated.groupby(['Location Number','Customer WO#','replaced_colocated_site'])['Total'].sum().reset_index()
gsf_kbs_subtotal.head()

#print('here')
#print(gsf_kbs_subtotal)

# Create table for getting PO info for payee central
# Define the column names
columns = [
    "Item",
    "Location Number",
    "Customer WO#",
    "DPE Master.Current Quarter PO.2",
    "PO Line",
    "PO Available Amount",
    "Description",
    "Line Description"
]

# Create an empty DataFrame
get_info_from_payee_central = pd.DataFrame(columns=columns)

# Replaced_colocated_site
get_info_from_payee_central['Location Number'] = gsf_kbs_invoices_w_colocated['replaced_colocated_site']
get_info_from_payee_central['Customer WO#'] = gsf_kbs_invoices_w_colocated['Customer WO#']
get_info_from_payee_central = get_info_from_payee_central.drop_duplicates(subset=['Customer WO#'])

#print(gsf_kbs_subtotal)
# Merge dataframe against itself to add subtotal from original and colocated location numbers
gsf_kbs_subtotal2 = gsf_kbs_subtotal.merge(gsf_kbs_subtotal, how='left', left_on=['Location Number','Customer WO#'], right_on=['replaced_colocated_site','Customer WO#'])
#print(gsf_kbs_subtotal2)

# Add subtotal for both colocated sites to compare against available amount
#print(gsf_kbs_subtotal2.columns)
gsf_kbs_subtotal2['Total_sum'] = gsf_kbs_subtotal2.apply(
    lambda row: row['Total_x'] + row['Total_y']
    if row['replaced_colocated_site_x'] != row['Location Number_x'] and pd.notna(row['Total_y'])
    else row['Total_x'],  # Default value if condition is not met
    axis=1
)

# Fix headers
gsf_kbs_subtotal2=gsf_kbs_subtotal2.drop(['replaced_colocated_site_y'], axis=1)

"""### Export tables as csv"""

gsf_kbs_invoices_w_colocated.to_csv(r'C:\Users\it-prod-bot4\Documents\Amazon GSF\02 Outputs\1-gsf_kbs_invoices_w_colocated.csv', index=False)
gsf_kbs_subtotal2.to_csv(r'C:\Users\it-prod-bot4\Documents\Amazon GSF\02 Outputs\1-gsf_kbs_subtotal.csv', index=False)
get_info_from_payee_central.to_excel(r'C:\Users\it-prod-bot4\Documents\Amazon GSF\02 Outputs\1-get_info_from_payee_central.xlsx', index=False, sheet_name="Sheet1")
