import os
from threading import Thread
from _thread import *
from json import JSONDecodeError
from bs4 import UnicodeDammit
import pandas as pd
import os
import time
import requests
import pandas as pd
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from msedge.selenium_tools import Edge, EdgeOptions
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.common.action_chains import ActionChains
from dotenv import load_dotenv
import re
import csv

load_dotenv()
USER=os.getenv("USER")
PASSWORD=os.getenv("PASSWORD")


# # unicode_chars = 'å∫ç'
EDGE_DRIVER = r'msedgedriver.exe'

load_dotenv()
USER=os.getenv("USER")
PASSWORD=os.getenv("PASSWORD")


# # unicode_chars = 'å∫ç'
EDGE_DRIVER = r'msedgedriver.exe'


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
cols = ['pcnNumber',	'jpcnAnalysisStgDesc',	'supName',	'supPcnId',	'supContactPhone',	'pcnEffectDt',	'changeReason',	'changeDescription',	'pcnCompliance','pcnComplianceDesc',	'ceRecommendation',	'ceRecommendationDesc',
                    	'ceQualReportComment', 'coreCeRecommendationComment',	'ceInitialAssessDt',	'jpn',	'mpn', 'PR',	'deviation',	'ChangeOrder',	'QPET',	'Assy']

final_df = final_df[cols]
with pd.ExcelWriter(r'fold\Output.xlsx', mode='w', engine='xlsxwriter') as writer:  
    final_df.to_excel(writer, sheet_name='Sheet_1')

# browser.close()

def find_mpn(text):
    res = []
    temp = text.split()
    for idx in temp:
        if any(chr.isalpha() for chr in idx) and any(chr.isdigit() for chr in idx) and not any(chr.contains(":") for chr in idx):
            res.append(idx)
    #num = re.findall(r'\b^[A-Z][A-Za-z0-9-]*$\b',text)
    return ",".join(res)


s = requests.Session()
# Set correct user agent
AAD_AUTHORITY_HOST_URI = r'https://login.microsoftonline.com'
AAD_TENANT_ID = r'bea78b3c-4cdb-4130-854a-1d193232e5f4'
browserOptions = EdgeOptions()
browserOptions.add_experimental_option("prefs", {
    "download.default_directory": r'C:\Day-to-Day\MY_WORK_OTHER\Sele\downloadsFiles\\',
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})
browserOptions.use_chromium = True
browserOptions.add_argument("-inprivate")
browserOptions.add_argument("--disable-notifications")
browserOptions.add_argument("--no-sandbox")
browserOptions.add_argument("--disable-software-rasterizer")
browserOptions.add_argument("--disable-gpu")
browser = Edge(executable_path=EDGE_DRIVER, options=browserOptions)
selenium_user_agent = browser.execute_script("return navigator.userAgent;")
s.headers.update({"user-agent": selenium_user_agent})

browser.get(r'https://iam-signin.juniper.net/app/juniper_agileplm_1/exk1mv21lsoJQuzDl0h8/sso/saml')

# main_page = browser.current_window_handle

wait = WebDriverWait(browser, 300)
waitLoginURL = r"https://iam-signin.juniper.net"

wait.until(EC.url_contains(waitLoginURL))
nameWait = WebDriverWait(browser, 20)

nameWait.until(EC.visibility_of_any_elements_located((By.ID, r"idp-discovery-username")))
browser.find_element(By.NAME, r"username").send_keys(r"debayand@juniper.net")
browser.find_element(By.XPATH, r"//input[@type='submit' and @value='Next']").click()

wait2 = WebDriverWait(browser, 300)
ADaitLoginURL = AAD_AUTHORITY_HOST_URI + "/" + AAD_TENANT_ID + "/saml2"
wait2.until(EC.url_contains(ADaitLoginURL))
ADnameWait = WebDriverWait(browser, 20)
ADnameWait.until(EC.visibility_of_any_elements_located((By.NAME, r"loginfmt")))
browser.find_element(By.NAME, r"loginfmt").send_keys(USER)
browser.find_element(By.XPATH, r"//input[@type='submit' and @value='Next']").click()
wait2.until(EC.visibility_of_any_elements_located((By.NAME, r"passwd")))
browser.find_element(By.ID, r"i0118").send_keys(PASSWORD)
time.sleep(1)
browser.find_element(By.XPATH, r"//input[@type='submit' and @value='Sign in']").click()

time.sleep(3)
# https://agileplm.juniper.net/Agile/default/login-cms.jsp
browser.get(r'https://agileplm.juniper.net/Agile/PLMServlet')

browser.switch_to.window(browser.window_handles[1])
wait3 = WebDriverWait(browser, 300)
wait3LoginURL = r"https://agileplm.juniper.net/Agile/PLMServlet"
validState = ["Producton", "Prototype", "Preliminary_BOM", "Preliminary"]
wait3.until(EC.url_contains(wait3LoginURL))
jpnlist = ','.join(final_df.jpn.tolist()).split(',')

dict_jpn = {}
for jpn in jpnlist:
    dict_jpn[jpn]=[]
    WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH,"//input[@id='QUICKSEARCH_STRING']"))).clear()
    WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH,"//input[@id='QUICKSEARCH_STRING']"))).send_keys(jpn)
    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a#top_simpleSearch"))).click()
    time.sleep(4)
    try:
        Select(browser.find_element(By.XPATH, "//div[@id='quickClassOptions']/select[1]")).select_by_visible_text("Items".strip())
        time.sleep(1)
        Select(browser.find_element(By.XPATH, "//div[@id='quickClassOptions']/select[2]")).select_by_visible_text("Parts".strip())
        time.sleep(1)
       

    except:
        WebDriverWait(browser,20).until(EC.element_to_be_clickable((By.XPATH, "(//a[@class='image_link'])[1]")))
        time.sleep(3)
    else:
        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[@id='quickClassOptions']/a[1]"))).click()
        time.sleep(3)        

    
    finally:
        elem = WebDriverWait(browser,20).until(EC.element_to_be_clickable((By.XPATH,"//div[@id='tabsDiv']/ul/li[6]/a")))
        if elem.is_displayed() == True:
            elem.click()           
            try:
                browser.implicitly_wait(20)                    
                rowmpn = WebDriverWait(browser,10, ignored_exceptions=(NoSuchElementException,StaleElementReferenceException)).until(EC.presence_of_all_elements_located((By.XPATH, "//div[@class='GMBodyMid']/div/table/tbody/tr[@class='GMDataRow']")))
                whereuseds=[]              
                for i in range(len(rowmpn)):
                            # browser.implicitly_wait(30)
                    x= rowmpn[i].find_element(By.XPATH, ".//td[5]")
                    if str(x.text).startswith('P'):
                        time.sleep(2)
                        
                        wherusedeach = rowmpn[i].find_element(By.XPATH, ".//td[3]").text
                        time.sleep(1)
                        whereuseds.append(wherusedeach)
                        time.sleep(1)
                        print(i)                        
                    else:
                        break
                browser.find_element(By.CSS_SELECTOR,"a#top_home").click()
                time.sleep(2)                              
            except:
                break

    dict_jpn[jpn].append(whereuseds)
    print(dict_jpn)


pd.DataFrame(dict_jpn).to_csv('otpt.csv')

        
