import pandas as pd
import datetime as dt
import sys
from openpyxl import load_workbook


# xbr_timeperiod = sys.argv[-1]
# xbr_type = sys.argv[-2]
# data_file = sys.argv[-3]
# new_report = "WK" +  xbr_type + "_" + xbr_timeperiod + ".xlsx"

df = pd.read_csv('qbr.csv')
# df['SurveyRequest'] = df['Survey Request DateTime'].mask(pd.isnull, df['Survey Request DateTime'])
# df['SurveyRequest'].where(pd.notnull, df['Date site was turned over from BD to Ops'])

survey_request_columns_to_check = ["Survey Request DateTime", "Date site was turned over from BD to Ops"]
survey_complete_columns_to_check = ["Actual Survey Date", "On-site consultation actual date", "Survey Uploaded DateTime"]
asset_load_columns_to_check = ["Asset Load date","Date Trouble Ticket Created"]
installation_columns_columns_to_check = ["Activation Date","Installation Date"]
removal_columns_to_check = ["Removal Date","Scheduled Decommission date/time"]

df['SurveyRequest'] = df[survey_request_columns_to_check].bfill(axis=1).iloc[:, 0]
df['SurveyComplete'] = df[survey_complete_columns_to_check ].bfill(axis=1).iloc[:, 0]
df['AssetLoad'] = df[asset_load_columns_to_check].bfill(axis=1).iloc[:, 0]
df['Activation'] = df[installation_columns_columns_to_check].bfill(axis=1).iloc[:, 0]
df['Decomm'] = df[removal_columns_to_check].bfill(axis=1).iloc[:, 0]

def assign_region(country):
    if country in ["CA", "MX","US"]:
        return "NA"
    elif country in ["FR", "DE", "GB", "PT", "ES", "AT","IT"]:
        return "EU"
    elif country in ["JP","AU"]:
        return "APAC"
    else:
        return None
df['region'] = df['Country Code'].apply(assign_region)

def assign_program(recordType):
    if recordType in ["Locker Onboarding", "Locker Onboarding NA","Odin Onboarding EU"]:
        return "Locker"
    elif recordType in ["Location Onboarding", "Location"]:
        return "Apt Locker Pro"
    elif recordType in ["Dobby Onboarding", "Dobby Onboarding EU", "Dobby Onboarding NA"]:
        return "Apt Locker"
    else:
        return None
df['program'] = df['Case Record Type'].apply(assign_program)

tfile = "transform.csv"
csv_file = df.to_csv(header=['region', 'Country Code','program','Status','Survey Review outcome','SurveyRequest', 'SurveyComplete',  'AssetLoad', 'Activation',  'Decomm' ], columns=['region', 'Country Code','program', 'Survey Review outcome','Status','SurveyRequest', 'SurveyComplete',  'AssetLoad', 'Activation', 'Decomm' ])
df.to_csv(tfile, header=['region', 'Country Code','program','Status','Survey Review outcome','SurveyRequest', 'SurveyComplete',  'AssetLoad', 'Activation',  'Decomm' ], columns=['region', 'Country Code','program','Status','Survey Review outcome','SurveyRequest', 'SurveyComplete',  'AssetLoad', 'Activation', 'Decomm' ], encoding='utf-8', index=False)