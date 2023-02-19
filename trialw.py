import sys
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
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from dotenv import load_dotenv
import multiprocessing.dummy as mp 
load_dotenv()
USER=os.getenv("USER")
PASSWORD=os.getenv("PASSWORD")


# unicode_chars = 'å∫ç'
EDGE_DRIVER = r'msedgedriver.exe'
badChrs = r'<>:"/\|?*(#) '


def cleanFilename(sourcestring,  removestring =" #%:/,.\\[]<>*(?)"):
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
download_dir = r'C:\Day-to-Day\MY_WORK_OTHER\Sele\downloadsFiles'

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

wait3.until(EC.url_contains(wait3LoginURL))

browser.get(r'https://agileplm.juniper.net/Agile/PLMServlet')

def getFiles(dir):
    return [x for x in os.listdir(download_dir)]

jpnHome = EC.presence_of_element_located(By.CSS_SELECTOR, "div#tabsDiv")

try:
    for jpnVal in os.listdir(download_dir):
        browser.find_element(By.XPATH,"//input[@id='QUICKSEARCH_STRING']").clear()
        time.sleep(2)
        WebDriverWait(browser, 10).until(EC.visibility_of_any_elements_located((By.XPATH,"//input[@id='QUICKSEARCH_STRING']")))
        browser.find_element(By.XPATH, "//input[@id='QUICKSEARCH_STRING']").send_keys(jpnVal)
        time.sleep(2)
        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a#top_simpleSearch"))).click()
        time.sleep(4)
        try:                    
            Select(browser.find_element(By.XPATH, "//div[@id='quickClassOptions']/select[1]")).select_by_visible_text("Items".strip())
            time.sleep(1)
            Select(browser.find_element(By.XPATH, "//div[@id='quickClassOptions']/select[2]")).select_by_visible_text("Parts".strip())
            time.sleep(1)
            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[@id='quickClassOptions']/a[1]"))).click()
            time.sleep(5)  
            try:              
                WebDriverWait(browser,20).until(EC.element_to_be_clickable((By.XPATH,"//div[@id='tabsDiv']/ul/li[4]/a"))).click()
                all_mpns = WebDriverWait(browser,20).until(EC.visibility_of_all_elements_located((By.XPATH,"//tr[@class='GMDataRow']/td[5]/a[@class='image_link']")))
                for eachmpn in all_mpns:
                    try:
                        eachmpn.click()
                        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH,"//div[@id='tabsDiv']/ul/li[4]/a"))).click()
                        attlinks= WebDriverWait(browser, 10).until(EC.visibility_of_all_elements_located((By.XPATH,"//div[@class='GMBodyMid']/div/table[@class='GMSection']/tbody/tr[@class='GMDataRow']")))
                        if not len(attlinks)>0:
                            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH,"//a[@id='MSG_AddAttachment_2']"))).click()                                         
                            time.sleep(5)
                            pathname= os.path.join(download_dir,jpnVal)
                            for root, dirs, files in os.walk(pathname):                
                                for j in files:                 
                                    filename = os.path.join(root,j)
                                    time.sleep(7)
                                    inputfile = browser.find_element(By.XPATH,"//div/span/a[@id='browserFiles']/input[@type='file']")
                                    time.sleep(3)
                                    inputfile.send_keys(filename)
                                    time.sleep(5)
                                uploadbut = WebDriverWait(browser,20).until(EC.visibility_of_element_located((By.XPATH,"//a[@id='uploadFilesUM']")))
                                uploadbut.click()
                               
                            
                        jpnblink = WebDriverWait(browser,20).until(EC.text_to_be_present_in_element((By.XPATH, "//ul[@class='breadcrumbs']/li[2]"),jpnVal))
                        jpnblink.click()                        
                    except:
                        print("no mpn found for {}".format(jpnVal))            
            except:
                firstlink = WebDriverWait(browser,20).until(EC.element_to_be_clickable((By.XPATH, "(//a[@class='image_link'])[1]")))
                firstlink.click()
                try:
                    WebDriverWait(browser,20).until(EC.element_to_be_clickable((By.XPATH,"//div[@id='tabsDiv']/ul/li[4]/a"))).click()
                    all_mpns = WebDriverWait(browser,20).until(EC.visibility_of_all_elements_located((By.XPATH,"//tr[@class='GMDataRow']/td[5]/a[@class='image_link']")))
                    for eachmpn in all_mpns:
                        try:
                            eachmpn.click()
                            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH,"//div[@id='tabsDiv']/ul/li[4]/a"))).click()
                            attlinks= WebDriverWait(browser, 10).until(EC.visibility_of_all_elements_located((By.XPATH,"//div[@class='GMBodyMid']/div/table[@class='GMSection']/tbody/tr[@class='GMDataRow']")))
                            if not len(attlinks)>0:
                                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH,"//a[@id='MSG_AddAttachment_2']"))).click()                                         
                                time.sleep(5)
                                pathname= os.path.join(download_dir,jpnVal)
                                for root, dirs, files in os.walk(pathname):                
                                    for j in files:                 
                                        filename = os.path.join(root,j)
                                        time.sleep(7)
                                        inputfile = browser.find_element(By.XPATH,"//div/span/a[@id='browserFiles']/input[@type='file']")
                                        time.sleep(3)
                                        inputfile.send_keys(filename)
                                        time.sleep(5)
                                    uploadbut = WebDriverWait(browser,20).until(EC.visibility_of_element_located((By.XPATH,"//a[@id='uploadFilesUM']")))
                                    uploadbut.click()
                               
                            
                            jpnblink = WebDriverWait(browser,20).until(EC.text_to_be_present_in_element((By.XPATH, "//ul[@class='breadcrumbs']/li[2]"),jpnVal))
                            jpnblink.click()                        
                        except:
                            print("no mpn found for {}".format(jpnVal))
                except:                      
                    continue
        except: 
            WebDriverWait(browser,20).until(EC.element_to_be_clickable((By.XPATH,"//div[@id='tabsDiv']/ul/li[4]/a"))).click()
            all_mpns = WebDriverWait(browser,20).until(EC.visibility_of_all_elements_located((By.XPATH,"//tr[@class='GMDataRow']/td[5]/a[@class='image_link']")))
            for eachmpn in all_mpns:
                try:
                    eachmpn.click()
                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH,"//div[@id='tabsDiv']/ul/li[4]/a"))).click()
                    attlinks= WebDriverWait(browser, 10).until(EC.visibility_of_all_elements_located((By.XPATH,"//div[@class='GMBodyMid']/div/table[@class='GMSection']/tbody/tr[@class='GMDataRow']")))
                    if not len(attlinks)>0:
                        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH,"//a[@id='MSG_AddAttachment_2']"))).click()                                         
                        time.sleep(5)
                        pathname= os.path.join(download_dir,jpnVal)
                        for root, dirs, files in os.walk(pathname):                
                            for j in files:                 
                                filename = os.path.join(root,j)
                                time.sleep(7)
                                inputfile = browser.find_element(By.XPATH,"//div/span/a[@id='browserFiles']/input[@type='file']")
                                time.sleep(3)
                                inputfile.send_keys(filename)
                                time.sleep(5)
                            uploadbut = WebDriverWait(browser,20).until(EC.visibility_of_element_located((By.XPATH,"//a[@id='uploadFilesUM']")))
                            uploadbut.click()                         
                    jpnblink = WebDriverWait(browser,20).until(EC.text_to_be_present_in_element((By.XPATH, "//ul[@class='breadcrumbs']/li[2]"),jpnVal))
                    jpnblink.click()                        
                except:
                    print("no mpn found for {}".format(jpnVal))
except:
    pass

browser.switch_to.window(browser.window_handles[0])
browser.refresh()
time.sleep(2)
browser.switch_to.window(browser.window_handles[1])

