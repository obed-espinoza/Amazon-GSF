import pandas as pd
import numpy as np
from datetime import datetime, timedelta

folder_outputs = r'C:\Users\it-prod-bot4\Documents\Amazon GSF\02 Outputs'

# Created by the bot
consolidated_invoices = pd.read_csv(folder_outputs+'\\Consolidated Invoices.csv')
pc_template = pd.read_excel(folder_outputs + '\\5-pc_template.xlsx')
pc_template_from_pc = pd.read_excel(folder_outputs + '\\PC Template for Amazon Tracking Fields.xlsx')
invoices_to_consolidate = pd.read_excel(folder_outputs + '\\4-invoices_to_consolidate.xlsx')

pc_tempalte_w_consolidated = pc_template.merge(consolidated_invoices, how='left', left_on='Invoice Number',
right_on="Consolidated INV#")
pc_tempalte_w_consolidated_w_invoices = pc_tempalte_w_consolidated.merge(invoices_to_consolidate, how='left', on='Site - WO#')

tracking_fields = pd.DataFrame()
tracking_fields['NetSuite Invoice Number'] = pc_tempalte_w_consolidated_w_invoices['Document Number'].fillna(pc_tempalte_w_consolidated_w_invoices['Invoice Number'])
tracking_fields['Invoice Sent'] = 'Portal Auto Invoice'
tracking_fields['Invoice Sent Date'] = (datetime.now()).strftime('%m/%d/%Y')
tracking_fields['Portal Invoice Number'] = pc_tempalte_w_consolidated_w_invoices['Invoice Number']
tracking_fields = tracking_fields.drop_duplicates().reset_index(drop=True)

from_pc_template     = set(pc_template['Invoice Number'].dropna().unique())
from_pc_template_from_pc = set(pc_template_from_pc['Invoice Number'].dropna().unique())

same_count          = len(from_pc_template) == len(from_pc_template_from_pc)
all_template_in_pc  = from_pc_template.issubset(from_pc_template_from_pc)
all_pc_in_template  = from_pc_template_from_pc.issubset(from_pc_template)
invoices_match      = same_count and all_template_in_pc and all_pc_in_template

with open(folder_outputs + '\\invoices_match.txt', 'w') as f:
    f.write(str(invoices_match))
  
tracking_fields = tracking_fields.to_csv(folder_outputs + '\\6-tracking_fields.csv', index=False)

print(consolidated_invoices)
print(invoices_to_consolidate)
print(tracking_fields)
print(f"Invoices match: {invoices_match}")

