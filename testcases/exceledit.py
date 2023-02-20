import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import xlsxwriter


df = pd.read_excel('Highgate_week_20230215.xlsx')


print (df)



df1=df.reindex(columns= ['CensusID','ChainID','HotelName','DateTY','DateLY','PropSupTY','PropDemTY','PropRevTY','PropSupLY','PropDemLY','PropRevLY','CompSupTY','CompDemTY','CompRevTY','CompSupLY','CompDemLY','CompRevLY'])
print(df1)

writer = pd.ExcelWriter('Highgate_week_20230215.xlsx',engine = "xlsxwriter")
workbook  = writer.book

df1.to_excel(writer, index = False)
  
writer.close()



print("Removing")

# Creating a dataframe 
df = pd.read_excel('Highgate_week_20230215.xlsx')
column_list = df.columns
# Create a Pandas Excel writer using XlsxWriter engine.
writer = pd.ExcelWriter("Highgate_week_20230215.xlsx", engine='xlsxwriter')


df.to_excel(writer, sheet_name='Sheet1', startrow=1, header=False, index=False)

# Get workbook and worksheet objects
workbook  = writer.book
worksheet = writer.sheets['Sheet1']


for idx, val in enumerate(column_list):
    worksheet.write(0, idx, val)

writer.close()

print(pd.__version__)


#CensusID,ChainID,HotelName,DateTY,DateLY,PropSupTY,PropDemTY,PropRevTY,CompSupTY,CompDemTY,CompRevTY,PropSupLY,PropDemLY,PropRevLY,CompSupLY,CompDemLY,CompRevLY

#'CensusID','ChainID','HotelName','DateTY','DateLY','PropSupTY','PropDemTY','PropRevTY','PropSupLY','PropDemLY','PropRevLY','CompSupTY','CompDemTY','CompRevTY','CompSupLY','CompDemLY','CompRevLY'
