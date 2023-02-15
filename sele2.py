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
EDGE_DRIVER = r'msedgedriver.exe'
from dotenv import load_dotenv
load_dotenv()
USER=os.getenv("USER")
PASSWORD=os.getenv("PASSWORD")

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
browser.get(r'https://pcn.juniper.net/getAssessmentCountByUserRole?userId=debayand&role=ce')
browser.get(r'https://pcn.juniper.net/getContextCEAssessmentDetail?userId=debayand&role=ce&stageCode=completed&pageNo=0&pageSize=50')

respon = BeautifulSoup(browser.page_source, 'html.parser')
jresp = json.loads(respon.text)

with open ('content2.json', 'w') as f: 
    f.write(json.dumps(jresp))          # write entire content

#Subjct List to extract
pcnAllAttrs = jresp['pcnList']
# def unnest(d):
#     for val in d.values():
#         if isinstance(val, dict):
#             yield from unnest(val)
#         else:
#             yield val
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

pcnNo = []
JPNNo = []  
df = pd.DataFrame()                            # Final list of PCNs #  
for i in range(len(pcnAllAttrs)):
    eachpcnNo = pcnAllAttrs[i]['pcnNumber']
    # pcnNo.append(eachpcnNo)
    pcnUri = "https://pcn.juniper.net/getPCN" + "/" + eachpcnNo + "?userId=debayand&pcnNo" + "=" + eachpcnNo
    browser.get(pcnUri)
    time.sleep(3)
    # re.findall(r'\{(.*?)\}', browser.page_source)
    jpnrespon = BeautifulSoup(browser.page_source, 'lxml')
    jsonJpn = json.loads(jpnrespon.text)
                                                    #jsonJpn['pcnInfo']['jpnList'] #
    # jpnMpnInfo = list(unnest(jsonJpn['pcnInfo']))
    # jpnMPNDataFrame = pd.DataFrame.from_dict({jsonJpn['pcnInfo']}, orient='index').transpose()
    df = df.append(jsonJpn['pcnInfo'],ignore_index=True)

with pd.ExcelWriter(r'fold\Outcomp.xlsx', mode='w', engine='xlsxwriter') as writer:  
    df.to_excel(writer, sheet_name='Sheet_1')



df1 = pd.read_excel(r'fold\outputs.xlsx')
df2 = pd.read_excel(r'fold\Outcomp.xlsx')

df3 = pd.read_excel(r'fold\Outcomp.xlsx', usecols=list(df1.columns[df1.columns.isin(df2.columns)]))

final_df = pd.concat([df1,df3])

final_df.set_index(['pcnId'])
final_df['jpn'] = final_df.jpnList.apply(lambda x: ",".join(map(str,re.findall("'affectedJPN': '(.*?)', 'affectedMPN",x))))
final_df['mpn'] = final_df.jpnList.apply(lambda x: ",".join(map(str,re.findall("'affectedMPN': '(.*?)', 'jpnLifeCycl",x))))
final_df['PR'] = final_df.devList.apply(lambda x: ",".join(map(str,re.findall("'prId': '(.*?)', 'deviation",x))))
final_df['deviation'] = final_df.devList.apply(lambda x: ",".join(map(str,re.findall("'deviation': '(.*?)', 'cmODMFactorySite",x))))
final_df['ChangeOrder'] = final_df.devList.apply(lambda x: ",".join(map(str,re.findall("'mcoEco': '(.*?)', 'mcoEcoStatus",x))))
final_df['QPET'] = final_df.devList.apply(lambda x: ",".join(map(str,re.findall("'qpet': '(.*?)', 'manufacturerName",x))))
final_df['Assy'] = final_df.devList.apply(lambda x: ",".join(map(str,re.findall("'jnprAssembly': '(.*?)', 'jpn",x))))

final_df = final_df.drop(['Unnamed: 0', 'devList', 'jpnList','sqeType', 'sqeAnalysisStg', \
    'sqeStartDt', 'sqeCompletionDt', 'sqeRecommendation', 'sqeRecommendationDesc','ceRiskFlag',\
        'supSampleOwnerContact', 'pendingConcern', 'pendingConcernComment','pcnCloseDt', 'supplierECD', 'pcnCordinator', 'coreCeAnalysisStg'], axis=1)

with pd.ExcelWriter(r'fold\Output.xlsx', mode='w', engine='xlsxwriter') as writer:  
    final_df.to_excel(writer, sheet_name='Sheet_1')

browser.close()

def find_mpn(text):
    res = []
    temp = text.split()
    for idx in temp:
        if any(chr.isalpha() for chr in idx) and any(chr.isdigit() for chr in idx) and not any(chr.contains(":") for chr in idx):
            res.append(idx)
    #num = re.findall(r'\b^[A-Z][A-Za-z0-9-]*$\b',text)
    return ",".join(res)


from xlsx2html import xlsx2html
import io

with open(r"fold\Output.xlsx", "rb") as xlsx_file:
    out_file = io.StringIO()
    xlsx2html(xlsx_file, out_file, locale='en')
    out_file.seek(0)
    result_html = out_file.read()
    with open(r"fold\result_html.html", 'w',encoding="utf-8") as f:
        f.write(result_html)
        
