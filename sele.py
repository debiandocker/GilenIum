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
import xlsxwriter
import argparse
import json
import re
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

browser.get(r'https://pcn.juniper.net/ceassessment/assessments?stageCode=closure')


wait = WebDriverWait(browser, 300)
waitLoginURL = AAD_AUTHORITY_HOST_URI + "/" + AAD_TENANT_ID + "/saml2"

wait.until(EC.url_contains(waitLoginURL))
nameWait = WebDriverWait(browser, 20)
nameWait.until(EC.visibility_of_any_elements_located((By.NAME, 'loginfmt')))
browser.find_element(By.NAME, 'loginfmt').send_keys(r'debayand@juniper.net')
browser.find_element(By.XPATH, "//input[@type='submit' and @value='Next']").click()
wait.until(EC.visibility_of_any_elements_located((By.NAME, 'passwd')))
browser.find_element(By.ID, 'i0118').send_keys(r'Yqxv9DAP7pf8ZUu')
browser.find_element(By.XPATH, "//input[@type='submit' and @value='Sign in']").click()

time.sleep(3)
browser.get(r'https://pcn.juniper.net/getAssessmentCountByUserRole?userId=debayand&role=ce')
browser.get(r'https://pcn.juniper.net/getContextCEAssessmentDetail?userId=debayand&role=ce&stageCode=closure&pageNo=0&pageSize=20')

respon = BeautifulSoup(browser.page_source, 'html.parser')
jresp = json.loads(respon.text)

with open ('content.json', 'w') as f: 
    f.write(json.dumps(jresp))          # write entire content

#Subjct List to extract
pcnAllAttrs = jresp['pcnList']

def dict_list_to_df(df, col):
    """Return a Pandas dataframe based on a column that contains a list of JSON objects or dictionaries.
    Args:
        df (Pandas dataframe): The dataframe to be flattened.
        col (str): The name of the column that contains the JSON objects or dictionaries.
    Returns:
        Pandas dataframe: A new dataframe with the JSON objects or dictionaries expanded into columns.
    """

    rows = []
    for index, row in df[col].iteritems():
        for item in row:
            rows.append(item)
    df = pd.DataFrame(rows)
    return df

    
pcnNo = []
JPNNo = []
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

# print(df)

with pd.ExcelWriter(r'fold\outputs.xlsx', mode='w',engine='xlsxwriter') as writer:  
    df.to_excel(writer, sheet_name='Sheet_1')

