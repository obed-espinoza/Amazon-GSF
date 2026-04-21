import pandas as pd
import numpy as np
from datetime import datetime, timedelta

folder_outputs = r'C:\Users\it-prod-bot4\Documents\Amazon GSF\02 Outputs'

# Created by the bot
consolidated_invoices = pd.read_csv(folder_outputs+'\\Consolidated Invoices.csv')
pc_template = pd.read_excel(folder_outputs + '\\5-pc_template.xlsx')
invoices_to_consolidate = pd.read_excel(folder_outputs + '\\4-invoices_to_consolidate.xlsx')

pc_tempalte_w_consolidated = pc_template.merge(consolidated_invoices, how='left', left_on='Invoice Number', right_on="Consolidated INV#")
pc_tempalte_w_consolidated_w_invoices = pc_tempalte_w_consolidated.merge(invoices_to_consolidate, how='left', on='Site - WO#')

tracking_fields = pd.DataFrame()
tracking_fields['NetSuite Invoice Number'] = pc_tempalte_w_consolidated_w_invoices['Document Number'].fillna(pc_tempalte_w_consolidated_w_invoices['Invoice Number'])
tracking_fields['Invoice Sent'] = 'Portal Auto Invoice'
tracking_fields['Invoice Sent Date'] = (datetime.now()).strftime('%m/%d/%Y')
tracking_fields['Portal Invoice Number'] = pc_tempalte_w_consolidated_w_invoices['Invoice Number']
tracking_fields = tracking_fields.drop_duplicates().reset_index(drop=True)


tracking_fields = tracking_fields.to_csv(folder_outputs + '\\6-tracking_fields.csv', index=False)


print(consolidated_invoices)
print(invoices_to_consolidate)
print(tracking_fields)