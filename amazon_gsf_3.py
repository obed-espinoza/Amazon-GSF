import pandas as pd

folder_outputs = r'C:\Users\it-prod-bot4\Documents\Amazon GSF\02 Outputs'


gsf_kbs_subtotal = pd.read_csv(folder_outputs+'\\1-gsf_kbs_subtotal.csv')


po_lines_combined_output = pd.read_excel(folder_outputs +'\\2-PO Lines combined_output.xlsx')
print(po_lines_combined_output[['Site Name','Customer WO#']])

gsf_kbs_invoices_w_colocated = pd.read_csv(folder_outputs + '\\1-gsf_kbs_invoices_w_colocated.csv')
print(gsf_kbs_invoices_w_colocated)
print('Size'+str(gsf_kbs_invoices_w_colocated.reset_index(drop=True).shape[0]))

#print('gsf_kbs_invoices_w_colocated'+gsf_kbs_invoices_w_colocated.columns)
print('po_lines_combined_output'+po_lines_combined_output.columns)
print(po_lines_combined_output[['Customer WO#','Site Name', 'Available Amount']])
print('gsf_kbs_subtotal '+gsf_kbs_subtotal.columns)

# Location Number_x refers to location number
gsf_kbs_subtotal = gsf_kbs_subtotal.merge(po_lines_combined_output, how='left', left_on=['Location Number_x','Customer WO#'], right_on=['Site Name','Customer WO#'])
print('c')
print(gsf_kbs_subtotal.columns)
print(gsf_kbs_subtotal[['Location Number_x','Customer WO#','Available Amount']])
print(gsf_kbs_subtotal[['Location Number_y','Customer WO#','Available Amount']])

# Location Number_y refers to colocated location number
gsf_kbs_subtotal['Location Number_y'] = gsf_kbs_subtotal['replaced_colocated_site_x'].fillna('AAAA')
gsf_kbs_subtotal = gsf_kbs_subtotal.merge(po_lines_combined_output, how='left', left_on=['Location Number_y','Customer WO#'], right_on=['Site Name','Customer WO#'], suffixes=('_y','_z'))

print('gsf_kbs_subtotal')
print(gsf_kbs_subtotal.columns)
print(gsf_kbs_subtotal[['Location Number_y','Customer WO#']])
print(po_lines_combined_output)
print(gsf_kbs_subtotal)
print(gsf_kbs_subtotal[['Available Amount_y','Available Amount_z']])

#print(gsf_kbs_subtotal.columns)
# Keep only relevant columns
gsf_kbs_subtotal=gsf_kbs_subtotal[['Location Number_x', 'Customer WO#', 'replaced_colocated_site_x', 'Total_x', 'Total_y', 'Total_sum', 'Available Amount_y', 'Available Amount_z']]
# Rename columns
gsf_kbs_subtotal=gsf_kbs_subtotal.rename(columns={'Total_x':'Site_Total_x','Total_y':'Colocated_Total_y','Total_sum':'Total_sum_sum'},)
# Coalesce available amounts with colocated available amounts
gsf_kbs_subtotal['Available Amount'] = gsf_kbs_subtotal['Available Amount_y'].fillna(gsf_kbs_subtotal['Available Amount_z'])
# Drop unnecessary columns
gsf_kbs_subtotal = gsf_kbs_subtotal.drop(['Available Amount_y', 'Available Amount_z'], axis=1)
# Remove suffix from column names
gsf_kbs_subtotal.columns = gsf_kbs_subtotal.columns.str.rsplit('_', n=1).str[0]

print('Size'+str(gsf_kbs_invoices_w_colocated.reset_index(drop=True).shape[0]))
print(gsf_kbs_invoices_w_colocated.columns)
print(gsf_kbs_subtotal.columns)
#print(gsf_kbs_subtotal.columns)
gsf_kbs_invoices_w_colocated_w_available_amounts = gsf_kbs_invoices_w_colocated.merge(gsf_kbs_subtotal, how='left', left_on=['Site Code','Customer WO#'], right_on=['Location Number','Customer WO#'])
print('a')
print(gsf_kbs_invoices_w_colocated[['Site Code','Customer WO#']])
print(gsf_kbs_subtotal[['Location Number','Customer WO#','Available Amount']])

print('Size'+str(gsf_kbs_invoices_w_colocated_w_available_amounts.reset_index(drop=True).shape[0]))


#print(gsf_kbs_invoices_w_colocated_w_available_amounts.columns)
gsf_kbs_invoices_w_colocated_w_available_amounts['Total_<=_Available_Amount'] = gsf_kbs_invoices_w_colocated_w_available_amounts['Total_sum']<=gsf_kbs_invoices_w_colocated_w_available_amounts['Available Amount']
#print(gsf_kbs_invoices_w_colocated_w_available_amounts['Total_<=_Available_Amount'])

print('Size'+str(gsf_kbs_invoices_w_colocated_w_available_amounts.reset_index(drop=True).shape[0]))
print(gsf_kbs_invoices_w_colocated_w_available_amounts)


#print(gsf_kbs_invoices_w_colocated_w_available_amounts[['Site Code','Total','Available Amount','Total_<=_Available_Amount']].head(15))
exceptions_report = (gsf_kbs_invoices_w_colocated_w_available_amounts[gsf_kbs_invoices_w_colocated_w_available_amounts['Total_<=_Available_Amount']==False])
gsf_kbs_invoices_w_colocated_w_available_amounts = (gsf_kbs_invoices_w_colocated_w_available_amounts[gsf_kbs_invoices_w_colocated_w_available_amounts['Total_<=_Available_Amount']==True])

print('Size'+str(gsf_kbs_invoices_w_colocated_w_available_amounts.reset_index(drop=True).shape[0]))

gsf_kbs_invoices_w_colocated_w_available_amounts['Site - WO#'] = gsf_kbs_invoices_w_colocated_w_available_amounts['Site Code'] + ' - ' + gsf_kbs_invoices_w_colocated_w_available_amounts['Customer WO#']

exceptions_report.to_excel(folder_outputs + '\\3-exceptions_report.xlsx', index=False)
gsf_kbs_invoices_w_colocated_w_available_amounts.to_csv(folder_outputs+'\\3-gsf_kbs_invoices_w_colocated_w_available_amounts.csv')
gsf_kbs_subtotal.to_csv(folder_outputs+'\\3-gsf_kbs_subtotal_w_available_amount.csv')