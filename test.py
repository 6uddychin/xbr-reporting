import pandas as pd
import datetime as dt
from datetime import datetime, date, timedelta
import sys
from openpyxl import load_workbook

reportWeek = (datetime.today() - timedelta(days=10)).strftime("%V")
reportName = str("WK" + reportWeek + "_WBR.xlsx")

print(reportName)

df = pd.read_csv('qbr.csv')
df.to_excel("wk" + reportWeek + "_qbr.xlsx")