from json import JSONDecodeError
import pandas as pd
import os
import time
import requests
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from msedge.selenium_tools import Edge, EdgeOptions
# from selenium.webdriver import Chrome, ChromeOptions
from bs4 import BeautifulSoup
import openpyxl
from openpyxl import load_workbook
import xlsxwriter
import argparse
import json
import re
from urllib.parse import urljoin
from selenium.webdriver.support.ui import Select
from pivottablejs import pivot_ui
from shareplum import Site, Office365
from shareplum.site import Version
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
import itertools


from dotenv import load_dotenv
load_dotenv()
USER=os.getenv("USER")
PASSWORD=os.getenv("PASSWORD")

EDGE_DRIVER = r'C:\Day-to-Day\MY_WORK_OTHER\Sele\msedgedriver.exe'

s = requests.Session()
# Set correct user agent
AAD_AUTHORITY_HOST_URI = r'https://login.microsoftonline.com'
AAD_TENANT_ID = r'bea78b3c-4cdb-4130-854a-1d193232e5f4'
browserOptions = EdgeOptions()
browserOptions.add_experimental_option("prefs", {
    "download.default_directory": r'C:\Day-to-Day\MY_WORK_OTHER\Sele',
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})
browserOptions.use_chromium = True
browserOptions.add_argument("-inprivate")
browser = Edge(executable_path=EDGE_DRIVER, options=browserOptions)

selenium_user_agent = browser.execute_script("return navigator.userAgent;")
s.headers.update({"user-agent": selenium_user_agent})

browser.get(r'https://pcn.juniper.net')

# wait = WebDriverWait(browser, 300)
# waitLoginURL = AAD_AUTHORITY_HOST_URI + "/" + AAD_TENANT_ID + "/saml2"

# wait.until(EC.url_contains(waitLoginURL))
# nameWait = WebDriverWait(browser, 20)
# nameWait.until(EC.visibility_of_any_elements_located((By.NAME, 'loginfmt')))
# browser.find_element(By.NAME, 'loginfmt').send_keys(USER)
# browser.find_element(By.XPATH, "//input[@type='submit' and @value='Next']").click()
# wait.until(EC.visibility_of_any_elements_located((By.NAME, 'passwd')))
# browser.find_element(By.ID, 'i0118').send_keys(PASSWORD)
# browser.find_element(By.XPATH, "//input[@type='submit' and @value='Sign in']").click()

# import re # Importing the RegEx library

def find_or_restart(browser, text):
    src = browser.page_source # loading page source
    text_found = re.search(r'%s' % (text), src) # regex looking for text
    print("checking for text on page...")
    if text_found:
        return
    else:
        print("not found. restart test...")
        login(browser)

def login(browser):    
    wait = WebDriverWait(browser, 300)
    waitLoginURL = AAD_AUTHORITY_HOST_URI + "/" + AAD_TENANT_ID + "/saml2"
    wait.until(EC.url_contains(waitLoginURL))
    nameWait = WebDriverWait(browser, 20)
    nameWait.until(EC.visibility_of_any_elements_located((By.NAME, 'loginfmt')))
    browser.find_element(By.NAME, 'loginfmt').send_keys(USER)
    browser.find_element(By.XPATH, "//input[@type='submit' and @value='Next']").click()
    wait.until(EC.visibility_of_any_elements_located((By.NAME, 'passwd')))
    browser.find_element(By.ID, 'i0118').send_keys(PASSWORD)
    browser.find_element(By.XPATH, "//input[@type='submit' and @value='Sign in']").click()
    time.sleep(3)
    

login(browser)
find_or_restart(browser, "JPCN") 

browser.get(r'https://pcn.juniper.net/getAssessmentCountByUserRole?userId=debayand&role=ce')
browser.get(r'https://pcn.juniper.net/getContextCEAssessmentDetail?userId=debayand&role=ce&stageCode=closure&pageNo=0&pageSize=100')

respon = BeautifulSoup(browser.page_source, 'html.parser')

jresp = json.loads(respon.text)

with open (r'C:\Day-to-Day\MY_WORK_OTHER\Sele\content.json', 'w') as f: 
    f.write(json.dumps(jresp))          # write entire content

#Subjct List to extract
pcnAllAttrs = jresp['pcnList']

# def dict_list_to_df(df, col):
#     """Return a Pandas dataframe based on a column that contains a list of JSON objects or dictionaries.
#     Args:
#         df (Pandas dataframe): The dataframe to be flattened.
#         col (str): The name of the column that contains the JSON objects or dictionaries.
#     Returns:
#         Pandas dataframe: A new dataframe with the JSON objects or dictionaries expanded into columns.
#     """

#     rows = []
#     for index, row in df[col].iteritems():
#         for item in row:
#             rows.append(item)
#     df = pd.DataFrame(rows)
#     return df

    
# pcnNo = []
# JPNNo = []


df=pd.DataFrame()                              # Final list of PCNs #  
for i in range(len(pcnAllAttrs)):
    eachpcnNo = pcnAllAttrs[i]['pcnNumber']

    # pcnNo.append(eachpcnNo)
    pcnUri = "https://pcn.juniper.net/getPCN" + "/" + eachpcnNo + "?userId=debayand&pcnNo" + "=" + eachpcnNo
    browser.get(pcnUri)
    time.sleep(3)
    # re.findall(r'\{(.*?)\}', browser.page_source)
    jpnrespon = BeautifulSoup(browser.page_source, 'lxml')
    jsonJpn = json.loads(jpnrespon.text)
    df = df.append(jsonJpn['pcnInfo'],ignore_index=True)

# pcnnosfor = pd.read_excel("AllPCNNumbers.xlsx", sheet_name="Sheet1")
# jpcnist = pcnnosfor['JPCNS'].tolist()
# for i in range(len(jpcnist)):
    
#     pcnUri = "https://pcn.juniper.net/getPCN" + "/" + jpcnist[i] + "?userId=debayand&pcnNo" + "=" +jpcnist[i]
#     print(pcnUri)
#     browser.get(pcnUri)
#     time.sleep(3)
#     # re.findall(r'\{(.*?)\}', browser.page_source)
#     jpnrespon = BeautifulSoup(browser.page_source, 'lxml')
#     jsonJpn = json.loads(jpnrespon.text)
#     df = df.append(jsonJpn['pcnInfo'],ignore_index=True)

# df['jpnList'] = df.jpnList.astype(str).str.findall("(?<='affectedJPN':)[^,]+(?=,)|(?<='affectedMPN':)[^,]+(?=,)").explode().groupby(level=0).unique().str.join(',')
# df["devList"]=df['devList'].astype(str).str.findall(r"ECO-\d{5}|M\d{5}|SQ\d{5}|DEV-\d{5}|[A-Z0-9\s\-]{6,}").explode().groupby(level=0).unique().str.join(',')
# df3 = df.replace(to_replace=r"[\[\{\}\]]", value="", regex=True)
with pd.ExcelWriter(r'C:\Day-to-Day\MY_WORK_OTHER\Sele\fold\outputs.xlsx', mode='w',engine='xlsxwriter') as writer:  
    df.to_excel(writer, sheet_name='Sheet_1')
# df.drop(["pcnId","jpcnAnalysisStg","supId","supPcnId","changeNature","supplierStatus","supContactPhone","supOtherContactEmail","pcnIssueDt","pcnReceivedDt","ceInitialAssessDt", "coreCEAssessDt","ceClosureAssessDt",
#                "pcnType","pcnTypeDesc",	"pcnSource","pcnSourceDesc","quarter","changeNature","changeNatureDesc","pendingConcernDesc",
#                "respCE","respCoreCE",	"pcnStatus","pcnStatusDesc","pcnStatusComment","pcnCompliance","pcnComplianceDesc","ceRecommendation",
#                "ceRecommendationDesc","ceRecommendationComment","ceQualReportComment","ceQualReportAcceptStatus","ceQualReportAcceptStatusDesc",
#                "coreCeRecommendation","coreCeRecommendationDesc",	"coreCeRecommendationComment","coreCeQualReportAcceptStatus",
#                 "coreCeQualReportAcceptStatusDesc","coreCeQualRptComment","priorityType","priorityTypeDesc","pcnCloseComment","ceInitialAssessDt",
#                 "coreCEAssessDt","ceClosureAssessDt","partAnalysisStatus","supEscalationContactName","supEscalationContactEmail","routeToNPI",
#                 "attachmentList","qualRptReviewTemple","supInfoStatus","pcnCordinator","coreCeAnalysisStg","coreCeAnalysisStgDesc",
#                 "sqeAnalysisStg","sqeAnalysisStgDesc","sqeStartDt","sqeCompletionDt","sqeFactoryAudit","sqeFactoryAuditComments","sqeRecommendation",
#                 "sqeRecommendationDesc","ceRiskFlag","ceRiskComment","supSampleOwnerContact","pendingConcern",	"pendingConcernComment",
#                 "formFitFunctionImpact",	"pcnCloseDt",	"supplierECD",	"lastTimeBuyDt",	"lastTimeShipDt",	"whereUsedAnalysisStatus",
#                 "nonComplianceReason",	"nonComplianceReasonDesc",	"ceResolutionComments"],axis=1,inplace=True)

# with pd.ExcelWriter(r'fold\outputsqe.xlsx', mode='w',engine='xlsxwriter') as writer:  
#     df.to_excel(writer, sheet_name='Sheet_1')
time.sleep(3)
browser.get(r'https://pcn.juniper.net/getAssessmentCountByUserRole?userId=debayand&role=ce')

browser.get(r'https://pcn.juniper.net/getContextCEAssessmentDetail?userId=debayand&role=ce&stageCode=completed&pageNo=0&pageSize=100')

respon2 = BeautifulSoup(browser.page_source, 'html.parser')
jresp2 = json.loads(respon2.text)

with open ('content2.json', 'w') as f: 
    f.write(json.dumps(jresp2))          # write entire content

#Subjct List to extract
pcnAllAttrs2 = jresp2['pcnList']

df0 = pd.DataFrame()                            # Final list of PCNs #  
for i in range(len(pcnAllAttrs2)):
    eachpcnNo2 = pcnAllAttrs2[i]['pcnNumber']
    # pcnNo.append(eachpcnNo)
    pcnUri2 = "https://pcn.juniper.net/getPCN" + "/" + eachpcnNo2 + "?userId=debayand&pcnNo" + "=" + eachpcnNo2
    browser.get(pcnUri2)
    time.sleep(3)
    # re.findall(r'\{(.*?)\}', browser.page_source)
    jpnrespon2 = BeautifulSoup(browser.page_source, 'lxml')
    jsonJpn2 = json.loads(jpnrespon2.text)
                                                    #jsonJpn['pcnInfo']['jpnList'] #
    # jpnMpnInfo = list(unnest(jsonJpn['pcnInfo']))
    # jpnMPNDataFrame = pd.DataFrame.from_dict({jsonJpn['pcnInfo']}, orient='index').transpose()
    df0 = df0.append(jsonJpn2['pcnInfo'],ignore_index=True)

    

with pd.ExcelWriter(r'C:\Day-to-Day\MY_WORK_OTHER\Sele\fold\Outcomp.xlsx', mode='w', engine='xlsxwriter') as writer:  
    df0.to_excel(writer, sheet_name='Sheet_1')


# dfz = pd.DataFrame()  
# pcnnosfor = pd.read_excel("fold\DD_AIR.xlsx", sheet_name="Sheet1")
# jpcnist = pcnnosfor['JPCNS'].tolist()
# for i in range(len(jpcnist)):
    
#     pcnUri = "https://pcn.juniper.net/getPCN" + "/" + jpcnist[i] + "?userId=debayand&pcnNo" + "=" +jpcnist[i]
#     print(pcnUri)
#     browser.get(pcnUri)
#     time.sleep(3)
#     # re.findall(r'\{(.*?)\}', browser.page_source)
#     jpnrespon = BeautifulSoup(browser.page_source, 'lxml')
#     jsonJpn = json.loads(jpnrespon.text)
#     dfz = dfz.append(jsonJpn['pcnInfo'],ignore_index=True)
# with pd.ExcelWriter(r'fold\Outsss.xlsx', mode='w', engine='xlsxwriter') as writer:  
#     df0.to_excel(writer, sheet_name='Sheet_1')
df1 = pd.read_excel(r'C:\Day-to-Day\MY_WORK_OTHER\Sele\fold\outputs.xlsx')
df2 = pd.read_excel(r'C:\Day-to-Day\MY_WORK_OTHER\Sele\fold\Outcomp.xlsx')
# df4 = pd.read_excel(r'fold\Outsss.xlsx')

df3 = pd.read_excel(r'C:\Day-to-Day\MY_WORK_OTHER\Sele\fold\Outcomp.xlsx', usecols=list(df1.columns[df1.columns.isin(df2.columns)]))
# df5 = pd.read_excel(r'fold\Outsss.xlsx', usecols=list(df1.columns[df1.columns.isin(df4.columns)]))
final_df = pd.concat([df1,df3],ignore_index=True)
# final_df = pd.concat([final_df,df5],ignore_index=True)

final_df.set_index(['pcnId'])
final_df['JPN'] = final_df.jpnList.astype(str).str.findall("(?<='affectedJPN':)[^,]+(?=,)").explode().groupby(level=0).unique().str.join(',')
final_df['JPN'] = final_df['JPN'].astype(str).str.replace("[']", "", regex=True)
final_df['MPN'] = final_df.jpnList.astype(str).str.findall("(?<='affectedMPN':)[^,]+(?=,)").explode().groupby(level=0).unique().str.join(',')
final_df['MPN'] = final_df['MPN'].astype(str).str.replace("[']", "", regex=True)
final_df['MPN-AnnualDemand'] = final_df.jpnList.astype(str).str.findall("(?<='affectedMPN':)[^,]+(?=,)|(?<='annualDemand':)[^,]+(?=,)").explode().groupby(level=0).unique().str.join(',')
final_df['MPN-AnnualDemand'] = final_df['MPN-AnnualDemand'].astype(str).str.replace("[']", "", regex=True)

# final_df['PR']= final_df.devList.astype(str).str.findall("[1-9]{7}")
final_df['PR']= final_df.devList.astype(str).str.findall("(?<='prId':)[^,]+(?=,)")
# final_df['PR']= final_df['PR'].astype(str).str.replace("[']", "", regex=True)
# final_df['Deviation'] = final_df.devList.str.findall('DEV-\d{5}')
final_df['Deviation'] = final_df.devList.str.findall("(?<='deviation':)[^,]+(?=,)")
# final_df['Deviation'] = final_df['Deviation'].astype(str).str.replace("[']", "", regex=True)
final_df['Deviation Status'] = final_df.devList.str.findall("(?<='deviationStatus':)[^,]+(?=,)")
# final_df['Deviation Status'] = final_df['Deviation Status'].astype(str).str.replace("[']", "", regex=True)
final_df['HTR'] = final_df.devList.str.findall("(?<='htr':)[^,]+(?=,)").explode().groupby(level=0).unique().str.join(',')
# final_df['Change Order'] = final_df.devList.astype(str).str.findall("ECO-\d{5}|M\d{5}")
final_df['Change Order'] = final_df.devList.astype(str).str.findall("(?<='mcoEco':)[^,]+(?=,)").explode().groupby(level=0).unique().str.join(',')
final_df['Change Order'] = final_df['Change Order'].astype(str).str.replace("[']", "", regex=True)
final_df['Change Order Status'] = final_df.devList.astype(str).str.findall("(?<='mcoEcoStatus':)[^,]+(?=,)").explode().groupby(level=0).unique().str.join(',')
final_df['Change Order Status'] = final_df['Change Order Status'].astype(str).str.replace("[']", "", regex=True)
final_df['ODMSite'] = final_df.devList.astype(str).str.findall("(?<='cmODMFactorySite':)[^,]+(?=,)")
final_df['QPET'] = final_df.devList.astype(str).str.findall("(?<='qpet':)[^,]+(?=,)")
# final_df['ODMSite-QPET'] = final_df['ODMSite-QPET'].astype(str).str.replace("[']", "", regex=True)
final_df = final_df.drop(['Unnamed: 0','sqeType', 'sqeAnalysisStg', \
    'sqeStartDt', 'sqeCompletionDt', 'sqeRecommendation', 'sqeRecommendationDesc','ceRiskFlag',\
        'supSampleOwnerContact', 'pendingConcern', 'pendingConcernComment','pcnCloseDt', 'supplierECD', 'pcnCordinator', 'coreCeAnalysisStg'], axis=1)
cols = ['pcnNumber', 'supName', 'supPcnId',\
        'JPN', 'MPN','MPN-AnnualDemand', 'PR', 'Deviation',\
       'Deviation Status', 'Change Order', 'Change Order Status',\
       'ODMSite','QPET','pcnIssueDt', 'jpcnAnalysisStgDesc','changeReason', 'changeDescription', 'pcnComplianceDesc']


final_df = final_df[cols]

df1 = (final_df.apply(lambda x: list(itertools.zip_longest(x['PR'], x['Deviation'], x['Deviation Status'], x['ODMSite'],x['QPET'])), axis=1).explode().apply(lambda x: pd.Series(x, index=['PR', 'Deviation', 'Deviation Status', 'ODMSite','QPET'])).groupby(level=0).ffill())

full_df = final_df[['pcnNumber', 'supName', 'supPcnId','JPN', 'MPN','MPN-AnnualDemand', 'Change Order', 'Change Order Status','pcnIssueDt', 'jpcnAnalysisStgDesc','changeReason', 'changeDescription', 'pcnComplianceDesc']].join(df1)
full_df['PR']= full_df['PR'].astype(str).str.replace("[']", "", regex=True)
full_df['Deviation']= full_df['Deviation'].astype(str).str.replace("[']", "", regex=True)
full_df['Deviation Status']= full_df['Deviation Status'].astype(str).str.replace("[']", "", regex=True)
full_df['ODMSite']= full_df['ODMSite'].astype(str).str.replace("[']", "", regex=True)
full_df['QPET']= full_df['QPET'].astype(str).str.replace("[']", "", regex=True)

# # final_df.set_index('pcnNumber')
# (final_df.set_index(['pcnNumber']) 
#        .apply(lambda col: col.str.split(','))
#        .explode(['PR', 'Deviation','Deviation Status','ODMSite-QPET'])
#        .reset_index()
#        .reindex(final_df.columns, axis=1))

with pd.ExcelWriter(r'C:\Day-to-Day\MY_WORK_OTHER\Sele\fold\Output.xlsx', mode='w', engine='xlsxwriter') as writer:  
    full_df.to_excel(writer, sheet_name='Sheet_1', index=False, na_rep='NaN')
    wbook = writer.book
    for column in full_df:
        column_width = 22
        col_idx = full_df.columns.get_loc(column)
        text_wrap_format = wbook.add_format({'text_wrap': False,'valign': 'top', 'align':'center'})       
        writer.sheets['Sheet_1'].set_column(col_idx,col_idx,column_width, text_wrap_format)

    col_idx2 = full_df.columns.get_loc('ODMSite')
    writer.sheets['Sheet_1'].set_column(col_idx2,col_idx2,30)
    col_idx22 = full_df.columns.get_loc('QPET')
    writer.sheets['Sheet_1'].set_column(col_idx22,col_idx22,30)
    col_idx3 = final_df.columns.get_loc('changeReason')
    writer.sheets['Sheet_1'].set_column(col_idx,col_idx,30)
    col_idx4 = final_df.columns.get_loc('changeDescription')
    writer.sheets['Sheet_1'].set_column(col_idx,col_idx,30)


pivot_ui(full_df,outfile_path=r'C:\Day-to-Day\MY_WORK_OTHER\Sele\fold\pivottablejs.html',
    rows=['pcnNumber','supName','pcnIssueDt','JPN', 'MPN', 'PR', 'Deviation',
       'Deviation Status', 'Change Order', 'Change Order Status',
       'ODMSite','QPET'])



browser.close()
