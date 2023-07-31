import pandas as pd
from datetime import datetime, date, timedelta
import sys
import openpyxl as xl
from openpyxl import load_workbook
import time



activeDate = pd.to_datetime(datetime.today() + timedelta(days=10))
reportWeek = (datetime.today() - timedelta(days=10)).strftime("%V")
reportName = str("WK" + reportWeek + "PartnerOps_WBR.xlsx")
tempReport = str("wk" + reportWeek + "temp.xlsx")
dfName = ("wk" + reportWeek + "_qbr.xlsx")

qbrFile = "qbr_test.xlsx"


df1 = pd.read_csv('qbr.csv')
df1.to_excel(tempReport)

time.sleep(2)

df = pd.read_excel(tempReport)


# LIST OF FIELDS TO CHECK FOR CORRECT COLUMN
survey_request_columns_to_check = ["Survey Request DateTime", "Date site was turned over from BD to Ops"]
survey_complete_columns_to_check = ["Actual Survey Date", "On-site consultation actual date", "Survey Uploaded DateTime"]
asset_load_columns_to_check = ["Asset Load date","Date Trouble Ticket Created"]
installation_columns_columns_to_check = ["Activation Date","Installation Date"]
removal_columns_to_check = ["Removal Date","Scheduled Decommission date/time"]
active_pipeline_columns_to_check = ["Date/Time Closed", "Asset Load date","Date Trouble Ticket Created", "Activation Date","Installation Date"]


# GET THE FIRST NOT NULL VALUE IN THE FIELDS ABOVE
df['SurveyRequest'] = df[survey_request_columns_to_check].bfill(axis=1).iloc[:, 0]
df['SurveyComplete'] = df[survey_complete_columns_to_check ].bfill(axis=1).iloc[:, 0]
df['AssetLoad'] = df[asset_load_columns_to_check].bfill(axis=1).iloc[:, 0]
df['Activation'] = df[installation_columns_columns_to_check].bfill(axis=1).iloc[:, 0]
df['Decomm'] = df[removal_columns_to_check].bfill(axis=1).iloc[:, 0]

# CONVERT TIME/DATE FIELDS TO THE CORRECT DATE FORMATS FOR CYCLE TIMES
df['SurveyRequest'] = pd.to_datetime(df['SurveyRequest'])
df['SurveyComplete'] = pd.to_datetime(df['SurveyComplete'])
df['AssetLoad'] = pd.to_datetime(df['AssetLoad'])
df['Activation'] = pd.to_datetime(df['Activation'])

def badSurvey(date1,date2):
    if date1 > date2:
        return False
    else:
        return True

df['BadSurveyData'] = df.apply(lambda row: badSurvey(row['SurveyComplete'], row['SurveyRequest']), axis = 1 )
df['BadInstall'] = df.apply(lambda row: badSurvey(row['Activation'], row['AssetLoad']), axis = 1 )
df['BadEnd2End'] = df.apply(lambda row: badSurvey(row['Activation'], row['SurveyRequest']), axis = 1 )



print(df['BadSurveyData'])