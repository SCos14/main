# Right click and select "Create Console for Editor"
# To run code, highlight the desired block of code, and select "Shift+Enter". Practice by running the "Import" Statements below
# Import Morningstar and Data Analysis Libraries
import morningstar_data as md
import pandas as pd
import os
from datetime import date

# Assign Dataset and Search Criteria Ids to Variables Below
# To access personal Dataset and Search Criteria Ids from Morningstar Direct, use the following functions:
# md.direct.user_items.get_search_criteria()
# md.direct.user_items.get_data_sets()
# Once Ids are retrieved, assign them to the following variables and run the code (Ensure that the ids are stored in quotations e.g. "7107365")

# Dataset Id
data_set_id = "7107365"
# Search Criteria for Competitor Mutual Funds
comp_funds_id = "7105912"
# Search Criteria for AGF Mutual Funds
agf_funds_id = "7107602"
# Search Criteria for Competitor ETFs
comp_etf_id = "7105915"
# Search Criteria for AGF ETFs
agf_etf_id = "7109452"

# Run the remaining code (or simply use Ctrl+A, Shift+Enter) and watch the magic happen. Once executed, download the generated excel file denoted by todays date.
oef_df = md.direct.get_investment_data(investments= comp_funds_id,data_points= data_set_id)
agf_df = md.direct.get_investment_data(investments= agf_funds_id,data_points= data_set_id)
etf_df = md.direct.get_investment_data(investments= comp_etf_id,data_points= data_set_id)
agf_etf_df = md.direct.get_investment_data(investments= agf_etf_id,data_points= data_set_id)
    
df_list = []

for index, row in agf_df.iterrows():
    df_agf = agf_df.iloc[[index]]
    category = df_agf['Morningstar Category'][index]
    df_oef = oef_df[oef_df['Morningstar Category'] == category]
    df_merged = pd.concat([df_agf, df_oef]).reset_index(drop=True)
    df_list.append(df_merged)
for index, row in agf_etf_df.iterrows():
    df_agf = agf_etf_df.iloc[[index]]
    category = agf_etf_df['Morningstar Category'][index]
    df_etf = etf_df[etf_df['Morningstar Category'] == category]
    df_merged = pd.concat([df_agf, df_etf]).reset_index(drop=True)
    df_list.append(df_merged)
writer = pd.ExcelWriter('Competitive Intelligence Report - '+ str(date.today()) + '.xlsx', engine='xlsxwriter')
for df in df_list:
    sheet_name = df.loc[0, 'Id']
    df.to_excel(writer, sheet_name=sheet_name, index=False)
writer.save()