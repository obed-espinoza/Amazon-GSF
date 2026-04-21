import pandas as pd
import numpy as np

folder_outputs = r'C:\Users\it-prod-bot4\Documents\Amazon GSF\02 Outputs'
gsf_kbs_invoices_w_colocated_w_available_amounts= pd.read_csv(folder_outputs + '\\3-gsf_kbs_invoices_w_colocated_w_available_amounts.csv')

print(gsf_kbs_invoices_w_colocated_w_available_amounts)

row_counts_by_site = gsf_kbs_invoices_w_colocated_w_available_amounts.groupby(['Site Code','Customer WO#']).size().reset_index(
    name='row_count').sort_values(by='row_count', ascending=False).reset_index(drop=True)

print(gsf_kbs_invoices_w_colocated_w_available_amounts.columns)
invoices_to_consolidate_base = gsf_kbs_invoices_w_colocated_w_available_amounts[['Internal ID','Site Code','Customer WO#','Site - WO#','Document Number','Subtotal','Total','replaced_colocated_site_y']]
invoices_to_consolidate = invoices_to_consolidate_base.merge(row_counts_by_site, how='left', left_on=['Site Code','Customer WO#'], right_on=['Site Code','Customer WO#'])


#invoices_to_consolidate['Site - WO#'] = invoices_to_consolidate['Site Code'] + ' - ' + invoices_to_consolidate['Customer WO#']
invoices_not_to_consolidate = invoices_to_consolidate[invoices_to_consolidate['row_count']==1]
invoices_to_consolidate = invoices_to_consolidate[invoices_to_consolidate['row_count']>1]
invoices_to_consolidate = invoices_to_consolidate.sort_values(by='Customer WO#', ascending=True)
invoices_to_consolidate = invoices_to_consolidate.drop_duplicates()

unique_invoices_to_consolidate = invoices_to_consolidate.groupby(['Site Code', 'Customer WO#','Site - WO#'])[['Total', 'Subtotal']].sum().reset_index()
#unique_invoices_to_consolidate = invoices_to_consolidate.drop_duplicates(subset=['Site Code','Customer WO#'])
#invoices_to_consolidate = invoices_to_consolidate[invoices_to_consolidate['Site Code'].isin(['AVP9','HVP1'])]
print(invoices_to_consolidate)
print(unique_invoices_to_consolidate)

unique_invoices_to_consolidate = unique_invoices_to_consolidate.to_excel( folder_outputs+'\\4-unique_invoices_to_consolidate.xlsx', index=False)
invoices_to_consolidate = invoices_to_consolidate.to_excel( folder_outputs+'\\4-invoices_to_consolidate.xlsx', index=False)




