from json import JSONDecodeError
import pandas as pd
# from selenium.webdriver import Chrome, ChromeOptions
from bs4 import BeautifulSoup
import openpyxl
from openpyxl import load_workbook
import xlsxwriter
import argparse
import json
import re
from urllib.parse import urljoin

from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils.dataframe import dataframe_to_rows


df1 = pd.read_excel(r'C:\Day-to-Day\MY_WORK_OTHER\Sele\fold\outputs.xlsx')
df2 = pd.read_excel(r'C:\Day-to-Day\MY_WORK_OTHER\Sele\fold\Outcomp.xlsx')
# df4 = pd.read_excel(r'fold\Outsss.xlsx')

df3 = pd.read_excel(r'C:\Day-to-Day\MY_WORK_OTHER\Sele\fold\Outcomp.xlsx', usecols=list(df1.columns[df1.columns.isin(df2.columns)]))
# df5 = pd.read_excel(r'fold\Outsss.xlsx', usecols=list(df1.columns[df1.columns.isin(df4.columns)]))
final_df = pd.concat([df1,df3],ignore_index=True)
# final_df = pd.concat([final_df,df5],ignore_index=True)

final_df.set_index(['pcnId'])
final_df['JPN'] = final_df.jpnList.astype(str).str.findall("(?<='affectedJPN':)[^,]+(?=,)")
# final_df['JPN'] = final_df['JPN'].astype(str).str.replace("[']", "", regex=True)
final_df['MPN'] = final_df.jpnList.astype(str).str.findall("(?<='affectedMPN':)[^,]+(?=,)")
# final_df['MPN'] = final_df['MPN'].astype(str).str.replace("[']", "", regex=True)
final_df['MPN-AnnualDemand'] = final_df.jpnList.astype(str).str.findall("(?<='affectedMPN':)[^,]+(?=,)|(?<='annualDemand':)[^,]+(?=,)")
# final_df['MPN-AnnualDemand'] = final_df['MPN-AnnualDemand'].astype(str).str.replace("[']", "", regex=True)

# final_df['PR']= final_df.devList.astype(str).str.findall("[1-9]{7}")
final_df['PR']= final_df.devList.astype(str).str.findall("(?<='prId':)[^,]+(?=,)")
# final_df['PR']= final_df['PR'].astype(str).str.replace("[']", "", regex=True)
# final_df['Deviation'] = final_df.devList.str.findall('DEV-\d{5}')
final_df['Deviation'] = final_df.devList.str.findall("(?<='deviation':)[^,]+(?=,)")
# final_df['Deviation'] = final_df['Deviation'].astype(str).str.replace("[']", "", regex=True)
final_df['Deviation Status'] = final_df.devList.str.findall("(?<='deviationStatus':)[^,]+(?=,)")
# final_df['Deviation Status'] = final_df['Deviation Status'].astype(str).str.replace("[']", "", regex=True)
final_df['HTR'] = final_df.devList.str.findall("(?<='htr':)[^,]+(?=,)")
# final_df['Change Order'] = final_df.devList.astype(str).str.findall("ECO-\d{5}|M\d{5}")
final_df['Change Order'] = final_df.devList.astype(str).str.findall("(?<='mcoEco':)[^,]+(?=,)")
# final_df['Change Order'] = final_df['Change Order'].astype(str).str.replace("[']", "", regex=True)
final_df['Change Order Status'] = final_df.devList.astype(str).str.findall("(?<='mcoEcoStatus':)[^,]+(?=,)")
# final_df['Change Order Status'] = final_df['Change Order Status'].astype(str).str.replace("[']", "", regex=True)
final_df['ODMSite-QPET'] = final_df.devList.astype(str).str.findall("(?<='cmODMFactorySite':)[^,]+(?=,)|(?<='qpet':)[^,]+(?=,)")
# final_df['ODMSite-QPET'] = final_df['ODMSite-QPET'].astype(str).str.replace("[']", "", regex=True)
final_df = final_df.drop(['Unnamed: 0','sqeType', 'sqeAnalysisStg', \
    'sqeStartDt', 'sqeCompletionDt', 'sqeRecommendation', 'sqeRecommendationDesc','ceRiskFlag',\
        'supSampleOwnerContact', 'pendingConcern', 'pendingConcernComment','pcnCloseDt', 'supplierECD', 'pcnCordinator', 'coreCeAnalysisStg'], axis=1)
cols = ['pcnNumber', 'supName', 'supPcnId',\
        'JPN', 'MPN','MPN-AnnualDemand', 'PR', 'Deviation',\
       'Deviation Status', 'Change Order', 'Change Order Status',\
       'ODMSite-QPET','pcnIssueDt', 'jpcnAnalysisStgDesc','changeReason', 'changeDescription', 'pcnComplianceDesc']


final_df = final_df[cols]
