import os
from _thread import *
from bs4 import BeautifulSoup
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
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
import openpyxl
from openpyxl import load_workbook
import xlsxwriter
import shutil

# unicode_chars = 'å∫ç'
EDGE_DRIVER = r'msedgedriver.exe'
from dotenv import load_dotenv
load_dotenv()
USER=os.getenv("USER")
PASSWORD=os.getenv("PASSWORD")

def cleanFilename(sourcestring,  removestring =" #%:/,.@$!~\\[]<>*(?)"):
    return ''.join([c for c in sourcestring if c not in removestring])

def filesOnly(path):
    for file in os.listdir(path):
        if os.path.isfile(os.path.join(path, file)):
            yield file

s = requests.Session()
# Set correct user agent
AAD_AUTHORITY_HOST_URI = r'https://login.microsoftonline.com'
AAD_TENANT_ID = r'bea78b3c-4cdb-4130-854a-1d193232e5f4'
browserOptions = EdgeOptions()
browserOptions.add_experimental_option("prefs", {
    "download.default_directory": r"C:\Day-to-Day\Work Started\PCN\\",
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
browserOptions.add_argument("--disable-extensions")
browserOptions.add_argument("--disable-popup-blocking")

browser = Edge(executable_path=EDGE_DRIVER, options=browserOptions)


selenium_user_agent = browser.execute_script("return navigator.userAgent;")
s.headers.update({"user-agent": selenium_user_agent})
download_dir = r"C:\Day-to-Day\Work Started\PCN"

browser.get(r'https://pcn.juniper.net')

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
browser.get(r'https://pcn.juniper.net/ceassessment/assessments?stageCode=initial')
browser.implicitly_wait(10)
elem = WebDriverWait(browser,20).until(EC.presence_of_all_elements_located((By.XPATH, "//td[2]")))
browser.implicitly_wait(10)
pcns = []
for i in elem:
    pcns.append(i.text)

# browser.current_window_handle
for element in pcns:
    pcnurl = "https://pcn.juniper.net/ceassessment" + "/"+str(element)+"/"+"detail"
    browser.get(pcnurl)    
    attachlinks = WebDriverWait(browser,20).until(EC.presence_of_all_elements_located((By.XPATH, "//img[@title='download']")))
    time.sleep(3)
    for i in range(len(attachlinks)):
        attachlinks[i].click()
        time.sleep(5)                    
        x1=0
        while x1==0:
            count=0
            li = filesOnly(download_dir)
            for x1 in li:
                if x1.endswith(".crdownload"):
                    count = count+1        
            if count==0:
                    x1=1
            else:
                x1=0
        for file_name in filesOnly(download_dir):
            filename,ext = os.path.splitext(file_name)  
            file_name.replace(filename,cleanFilename(filename))
            pathname = os.path.join(download_dir,element)
            if not os.path.exists(pathname):
                os.mkdir(pathname)
            shutil.move(os.path.join(download_dir,file_name),os.path.join(pathname,file_name))



# browser.close()
# browser.quit()
        
