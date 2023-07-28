import pandas as pd
from openpyxl import load_workbook

template_file = "template.xlsx" 
new_file = "test.xlsx"

book = load_workbook(template_file)
writer = pd.ExcelWriter(new_file, engine = 'openpyxl') 
writer.book = book

pd.read_csv('adjust.csv').to_excel(writer, sheet_name = 'Data', index = False)

# Saving the file as 'test.xlsx'
writer.save()
writer.close()