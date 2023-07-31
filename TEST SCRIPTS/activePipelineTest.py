import pandas as pd
import datetime as dt
from datetime import datetime, date, timedelta
import sys
from openpyxl import load_workbook


df = pd.read_csv('qbr.csv')


activeDate = pd.to_datetime(datetime.today() + timedelta(days=10))


survey_request_columns_to_check = ["Survey Request DateTime", "Date site was turned over from BD to Ops"]
survey_complete_columns_to_check = ["Actual Survey Date", "On-site consultation actual date", "Survey Uploaded DateTime"]
asset_load_columns_to_check = ["Asset Load date","Date Trouble Ticket Created"]
installation_columns_columns_to_check = ["Activation Date","Installation Date"]
removal_columns_to_check = ["Removal Date","Scheduled Decommission date/time"]
active_pipeline_columns_to_check = ["Date/Time Closed", "Asset Load date","Date Trouble Ticket Created", "Activation Date","Installation Date"]

df['SurveyRequest'] = df[survey_request_columns_to_check].bfill(axis=1).iloc[:, 0]
df['SurveyComplete'] = df[survey_complete_columns_to_check ].bfill(axis=1).iloc[:, 0]
df['AssetLoad'] = df[asset_load_columns_to_check].bfill(axis=1).iloc[:, 0]
df['Activation'] = df[installation_columns_columns_to_check].bfill(axis=1).iloc[:, 0]
df['Decomm'] = df[removal_columns_to_check].bfill(axis=1).iloc[:, 0]

def activeFunc (x):
    if x == "NaN":
        activeDate = date.today() + timedelta(days=10)
        return activeDate
    else:
        df[active_pipeline_columns_to_check].bfill(axis=1).iloc[:, 0]

df['Active'] = df[active_pipeline_columns_to_check].bfill(axis=1).iloc[:, 0]


df['ActivePipeline'] = df['Active'].mask(df['Active'].isna(), activeDate)

df.to_excel("testActive.xlsx", header=['Active', 'ActivePipeline'], columns=['Active', 'ActivePipeline'])
