from json import JSONDecodeError
import pandas as pd
import os
import time
import requests
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from msedge.selenium_tools import Edge, EdgeOptions
# from selenium.webdriver import Chrome, ChromeOptions
from bs4 import BeautifulSoup
import openpyxl
from openpyxl import load_workbook
import argparse
import json
import re
from urllib.parse import urljoin
EDGE_DRIVER = r'msedgedriver.exe'
from dotenv import load_dotenv
load_dotenv()
USER=os.getenv("USER")
PASSWORD=os.getenv("PASSWORD")

load_dotenv()
USER=os.getenv("USER")
PASSWORD=os.getenv("PASSWORD")


# # unicode_chars = 'å∫ç'
EDGE_DRIVER = r'msedgedriver.exe'

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

browser.get(r'https://pcn.juniper.net/ceassessment/assessments?stageCode=completed')

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
# WebDriverWait(browser,10).until(EC.url_matches((r'https://pcn.juniper.net/ceassessment/assessments?stageCode=initial')))

# browser.get(r'https://pcn.juniper.net/getAssessmentCountByUserRole?userId=debayand&role=ce')

# browser.get(r'https://pcn.juniper.net/getContextCEAssessmentDetail?userId=debayand&role=ce&stageCode=completed&pageNo=0&pageSize=50')

# respon = BeautifulSoup(browser.page_source, 'html.parser')
# jresp = json.loads(respon.text)

# with open ('content2.json', 'w') as f: 
#     f.write(json.dumps(jresp))          # write entire content

# #Subjct List to extract
# pcnAllAttrs = jresp['pcnList']


# pcnNo = []
# JPNNo = []  
# df = pd.DataFrame()                            # Final list of PCNs #  
# for i in range(len(pcnAllAttrs)):
#     eachpcnNo = pcnAllAttrs[i]['pcnNumber']
#     # pcnNo.append(eachpcnNo)
#     pcnUri = "https://pcn.juniper.net/getPCN" + "/" + eachpcnNo + "?userId=debayand&pcnNo" + "=" + eachpcnNo
#     browser.get(pcnUri)
#     time.sleep(3)
#     # re.findall(r'\{(.*?)\}', browser.page_source)
#     jpnrespon = BeautifulSoup(browser.page_source, 'lxml')
#     jsonJpn = json.loads(jpnrespon.text)
#                                                     #jsonJpn['pcnInfo']['jpnList'] #
#     # jpnMpnInfo = list(unnest(jsonJpn['pcnInfo']))
#     # jpnMPNDataFrame = pd.DataFrame.from_dict({jsonJpn['pcnInfo']}, orient='index').transpose()
#     df = df.append(jsonJpn['pcnInfo'],ignore_index=True)

browser.get("https://gnats.juniper.net/web/default/all-my-prs")
respon = BeautifulSoup(browser.page_source, 'html5lib')
links = [node.get('href') for node in respon.find_all('a', attrs={'href': re.compile("-1$")})]

base_url = r'https://gnats.juniper.net'
full_links = [urljoin(base_url, link) for link in links]
tabs = ['summary_tab','description_tab','scope_tab','details_tab','assessment_tab','attachments_tab','external_tab',
        'fixinfo_tab','cases_tab','audit_tab','changelog_tab','cmevents_tab']

last_audit=[]
for i,link in enumerate(full_links):
    browser.get(full_links[i])
    time.sleep(5)
    browser.find_element(By.CSS_SELECTOR, 'a#audit-tab').click()
    time.sleep(10)
    audit_trails = []
    last_text = WebDriverWait(browser,30).until(EC.visibility_of_all_elements_located((By.XPATH, "(//div[@class='section-contents']/pre)[1]")))
    time.sleep(2)
    # trails = WebDriverWait(browser,30).until(EC.visibility_of_all_elements_located((By.XPATH, "//div[@class='section-contents']/pre")))
    # for trail in trails:
    #     new_text = '-------'+str(link)+trail.text+'\n\n'
    #     audit_trails.append(new_text)
    # print(audit_trails)
    last_audit.append(last_text)
    
dicts = {}
for i in range(len(links)):
    dicts[links[i]] = last_audit[i]

# with pd.ExcelWriter(r'fold\Outcomp.xlsx', mode='w', engine='xlsxwriter') as writer:  
#     df.to_excel(writer, sheet_name='Sheet_1')



df1 = pd.read_excel(r'fold\outputs.xlsx')
df2 = pd.read_excel(r'fold\Outcomp.xlsx')

df3 = pd.read_excel(r'fold\Outcomp.xlsx', usecols=list(df1.columns[df1.columns.isin(df2.columns)]))

# jpnList = [{'jpnID': 85070, 'affectedJPN': '740-073766', 'affectedMPN': 'FSH015-4C0G Rev05', 'newJPN': '740-073766', 'newMPN': 'FSH015-4C0G_Rev:06', 'jpnLifeCycl': 'Production', 
#   'mpnLifeCycl': 'Active', 'npiPrdNames': '', 'commcode': 74010, 'commGrp': 'Custom Power', 'commMgr': 'bengu', 'coreCE': 'rrasconh', 'pcnId': 101336, 'activeOtherAVLCount': '1', 
#   'annualDemand': '5202', 'manufactureName': 'ACBEL POLYTECH INC-TAIPEI CITY', 'commMgrName': 'Ben Gu', 'coreCEName': 'Rigoberto Rascon Hernandez', 'gcmCheckInventoryPosition': False, 
#   'npiList': [], 'contextCe': 'jbachal', 'contextCeName': 'Jatin Bachal', 'sqe': 'sguan', 'sqeName': 'Sam Guan'}]

# devList = [{'deviationId': 922, 'prId': '1655022', 'deviation': 'DEV-21582A', 'cmODMFactorySite': 'FLEXTRONICS-PENANG', 'jpn': '740-073766', 'mpn': 'FSH015-4C0G_Rev:06', 'priority': '10000', 
#             'htrMfgRptSubmit': 'yes', 'htr': '', 'mcoEco': 'ECO-54803', 'qpet': 'SQ01669', 'deviationStatus': 'released', 'manufacturerName': 'ACBEL POLYTECH INC-TAIPEI CITY', 'qpetStatus': 'Open', 
#             'qpetCurrentActivity': 'Deviation'}, {'deviationId': 923, 'prId': '1655022', 'deviation': 'DEV-21601A', 'cmODMFactorySite': 'FOXCONN JUAREZ', 'jpn': '740-073766', 'mpn': 'FSH015-4C0G_Rev:06', 
#                                                   'mcoEco': 'ECO-54803', 'qpet': 'SQ01669', 'deviationStatus': 'released', 'manufacturerName': 'ACBEL POLYTECH INC-TAIPEI CITY', 'qpetStatus': 'Open', 
#                                                   'qpetCurrentActivity': 'Deviation'}]

# df3['jpnList'] = df3.jpnList.astype(str).str.findall("(?<='affectedJPN':)[^,]+(?=,)|(?<='affectedMPN':)[^,]+(?=,)").explode().groupby(level=0).unique().str.join(',')

final_df = pd.concat([df1,df3],ignore_index=True)

final_df.set_index(['pcnId'])
final_df['JPN-MPN-AnnualDemand'] = final_df.jpnList.astype(str).str.findall("(?<='affectedJPN':)[^,]+(?=,)|(?<='affectedMPN':)[^,]+(?=,)|(?<='annualDemand':)[^,]+(?=,)").explode().groupby(level=0).unique().str.join(',')

final_df['PR']= final_df.devList.astype(str).str.findall("\d{7}").explode().groupby(level=0).unique().str.join(',')
final_df['Deviation'] = final_df.devList.str.findall('DEV-\d{5}').explode().groupby(level=0).unique().str.join(',')
final_df['Change Order'] = final_df.devList.astype(str).str.findall("ECO-\d{5}|M\d{5}").explode().groupby(level=0).unique().str.join(',')
final_df['ODMSite-QPET'] = final_df.devList.astype(str).str.findall("(?<='cmODMFactorySite':)[^,]+(?=,)|(?<='qpet':)[^,]+(?=,)").explode().groupby(level=0).unique().str.join(',')

# final_df['PR-trails']=[]
# for pr in final_df['PR']:
#     link = "/web/default/"+pr+"-1"
#     full_link = urljoin(base_url, link)
#     browser.get(full_links)
#     time.sleep(10)
#     browser.find_element(By.CSS_SELECTOR, 'a#audit-tab').click()
#     time.sleep(10)    
#     trails = WebDriverWait(browser,30).until(EC.visibility_of_all_elements_located((By.XPATH, "//div[@class='section-contents']/pre")))
#     for trail in trails:
#         new_text = '-------'+str(link)+trail.text+'\n\n'
#         final_df['PR-trails'].append(new_text)
    

# # another type
# final_df["PR"]=final_df['devList'].astype(str).str.findall("[1-9][0-9]{6}").apply(lambda x: list(set(x))).str.join(',')


final_df = final_df.drop(['Unnamed: 0','sqeType', 'sqeAnalysisStg', \
    'sqeStartDt', 'sqeCompletionDt', 'sqeRecommendation', 'sqeRecommendationDesc','ceRiskFlag',\
        'supSampleOwnerContact', 'pendingConcern', 'pendingConcernComment','pcnCloseDt', 'supplierECD', 'pcnCordinator', 'coreCeAnalysisStg'], axis=1)
cols = ['pcnNumber', 'JPN-MPN-AnnualDemand', 'PR', 'Deviation', 'Change Order', 'ODMSite-QPET', 'jpcnAnalysisStgDesc','supName','supPcnId',	'supContactPhone',	'pcnEffectDt',	'changeReason',	'changeDescription','pcnCompliance','pcnComplianceDesc','ceRecommendation','ceRecommendationDesc',
                    	'ceQualReportComment', 'coreCeRecommendationComment',	'ceInitialAssessDt']

final_df = final_df[cols]
final_df["PR_lastTrail"] = final_df['PR'].map(dicts)

with pd.ExcelWriter(r'fold\Output.xlsx', mode='w', engine='xlsxwriter') as writer:  
    final_df.to_excel(writer, sheet_name='Sheet_1')


browser.close()

# def find_mpn(text):
#     res = []
#     temp = text.split()
#     for idx in temp:
#         if any(chr.isalpha() for chr in idx) and any(chr.isdigit() for chr in idx) and not any(chr.contains(":") for chr in idx):
#             res.append(idx)
#     #num = re.findall(r'\b^[A-Z][A-Za-z0-9-]*$\b',text)
#     return ",".join(res)


# from xlsx2html import xlsx2html
# import io

# with open(r"fold\Output.xlsx", "rb") as xlsx_file:
#     out_file = io.StringIO()
#     xlsx2html(xlsx_file, out_file, locale='en')
#     out_file.seek(0)
#     result_html = out_file.read()
#     with open(r"fold\result_html.html", 'w',encoding="utf-8") as f:
#         f.write(result_html)
