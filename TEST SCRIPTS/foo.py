import openpyxl
import pandas as pd
import time

#Prepare the spreadsheets to copy from and paste too.


#change NameOfTheSheet with the sheet name that includes the data
data = pd.read_excel("qbr_template.xlsx", sheet_name="Working")

time.sleep(20)

#save it to the 'NewSheet' in destfile
data.to_excel("newfile.xlsx", sheet_name='table')

df = pd.read_excel("newfile.xlsx")

df1 = df.drop(df.columns[1], axis = 1)
df2 = df.drop(df.index[1])

df1.to_excel("newfiledf1.xlsx", sheet_name='table', index=False)
df2.to_excel("newfiledf2.xlsx", sheet_name='table',index=False)