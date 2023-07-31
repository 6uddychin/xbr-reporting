import pandas as pd
import datetime as dt
import sys
from openpyxl import load_workbook


df = pd.read_csv('transform.csv')

df['SurveyRequest'] = pd.to_datetime(df['SurveyRequest'])
df['SurveyComplete'] = pd.to_datetime(df['SurveyComplete'])
df['AssetLoad'] = pd.to_datetime(df['AssetLoad'])
df['Activation'] = pd.to_datetime(df['Activation'])

def check_InstallCycleTime (row):
    if row['Activation'] > row['AssetLoad']: 
        return (row['Activation'] - row['AssetLoad']).days
    else:
        return ""

def check_e2e (row):
    if row['Activation'] > row['SurveyRequest']: 
        return (row['Activation'] - row['SurveyRequest']).days
    else:
        return ""
    
def check_wastedTrip(row):
    if row['Survey Review outcome'] == "Wasted Trip":
        return row['SurveyComplete']
    else:
        return ""

def check_SurveyCycleTime (row):
    if row['SurveyComplete'] > row['SurveyRequest']:
        return (row['SurveyComplete'] - row['SurveyRequest']).days
    else:
        return ""

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
df['SurveyCycleTime'] = df.apply(check_SurveyCycleTime, axis=1)
df['InstallCycleTime'] = df.apply(check_InstallCycleTime, axis=1)
df['e2eCycleTime'] = df.apply(check_e2e, axis=1)
df['WatedTrip'] = df.apply(check_wastedTrip, axis = 1)


tfile = "adjust.csv"
csv_file = df.to_csv(header=['region', 'Country Code','program','Status','SurveyRequest', 'SurveyComplete','WatedTrip',  'AssetLoad', 'Activation','Decomm' , 'SurveyCycleTime' ,'InstallCycleTime','e2eCycleTime'], columns=['region', 'Country Code','program','Status','SurveyRequest', 'SurveyComplete','WatedTrip',  'AssetLoad', 'Activation','Decomm' , 'SurveyCycleTime' ,'InstallCycleTime','e2eCycleTime'])
df.to_csv(tfile, header=['region', 'Country Code','program','Status','SurveyRequest', 'SurveyComplete','WatedTrip',  'AssetLoad', 'Activation','Decomm' , 'SurveyCycleTime' ,'InstallCycleTime','e2eCycleTime'], columns=['region', 'Country Code','program','Status','SurveyRequest', 'SurveyComplete','WatedTrip',  'AssetLoad', 'Activation','Decomm' , 'SurveyCycleTime' ,'InstallCycleTime','e2eCycleTime'], encoding='utf-8', index=False)
df.to_excel('transform.xlsx',header=['region', 'Country Code','program','Status','SurveyRequest', 'SurveyComplete','WatedTrip',  'AssetLoad', 'Activation','Decomm' , 'SurveyCycleTime' ,'InstallCycleTime','e2eCycleTime'], columns=['region', 'Country Code','program','Status','SurveyRequest', 'SurveyComplete','WatedTrip',  'AssetLoad', 'Activation','Decomm' , 'SurveyCycleTime' ,'InstallCycleTime','e2eCycleTime'], encoding='utf-8', index=False)