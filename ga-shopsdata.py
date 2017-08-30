
# coding: utf-8

# In[104]:

import pandas as pd
import numpy as np
import glob
import os

os.getcwd()

#brondata excelfile
inputbron='ga-shops/input/ga-shops.xlsx'
inputdata='ga-shops/input/data.xlsx'
xls= pd.read_excel(inputbron)

#sheetnames
names = pd.ExcelFile(inputbron, on_demand = True)
sheets = names.sheet_names[2:]
#sheets

#create dataframe for all sheets
all_data=pd.DataFrame()

#verzamel alle sheets in all data en skip de eerste 14 regels van elk sheet
for sheet in sheets:
    df= pd.read_excel(inputbron, sheet, skiprows=14, parse_cols = "A:C")
    df['titel'] = sheet
    all_data = all_data.append(df,ignore_index=True)
        


# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('ga-shops/output/data.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
all_data.to_excel(writer, sheet_name='products')

# Close the Pandas Excel writer and output the Excel file.
writer.save()
all_data.head(10)


# In[ ]:



