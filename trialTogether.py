import csv
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
import openpyxl
from openpyxl import load_workbook
import xlsxwriter
import argparse
import json
import re
import xml.etree.ElementTree as ElT
import tqdm
import shutil

# unicode_chars = 'å∫ç'
EDGE_DRIVER = r'msedgedriver.exe'

def count_files(direct):
    for root, dirs, files in os.walk(direct):
        return len(list(f for f in files if not f.endswith('.crdownload')))

class file_has_been_downloaded(object):
    def __init__(self, dir, number):
        self.dir = dir
        self.number = number

    def __call__(self, driver):
        print(count_files(dir), '->', self.number)
        return count_files(dir) > self.number

def download_file(url):
    local_filename = url.split('/')[-1]
    # NOTE the stream=True parameter below
    with requests.get(url, stream=True) as r:
        r.raise_for_status()
        with open(local_filename, 'wb') as f:
            for chunk in r.iter_content(chunk_size=8192): 
                # If you have chunk encoded response uncomment if
                # and set chunk_size parameter to None.
                #if chunk: 
                f.write(chunk)
    return local_filename
# def login_info():
#     with open("usernamesPasswords.txt", "r") as infile:
#         data = [line.rstrip().split(":") for line in infile]
#         username = data[1][0]
#         password = data[1][1]
#     return username, password

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
browser.find_element(By.NAME, r"loginfmt").send_keys(r'debayand@juniper.net')
browser.find_element(By.XPATH, r"//input[@type='submit' and @value='Next']").click()
wait2.until(EC.visibility_of_any_elements_located((By.NAME, r"passwd")))
browser.find_element(By.ID, r"i0118").send_keys(r'Yqxv9DAP7pf8ZUu')
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
i=0
for jpnVal in os.listdir(download_dir):

    browser.refresh()
    ADnameWait2 = WebDriverWait(browser, 5)
    ADnameWait2.until(EC.visibility_of_any_elements_located((By.XPATH,"//input[@id='QUICKSEARCH_STRING']")))
    browser.find_element(By.XPATH,"//input[@id='QUICKSEARCH_STRING']").clear()
    browser.find_element(By.XPATH, "//input[@id='QUICKSEARCH_STRING']").send_keys(jpnVal)
    time.sleep(2)
    browser.find_element(By.CSS_SELECTOR, "a#top_simpleSearch").click()
    time.sleep(2)   

    try:   
        Select(browser.find_element(By.XPATH, "//div[@id='quickClassOptions']/select[1]")).select_by_visible_text("Items".strip())
        time.sleep(1)
        Select(browser.find_element(By.XPATH, "//div[@id='quickClassOptions']/select[2]")).select_by_visible_text("Parts".strip())
        time.sleep(2)

        browser.find_element(By.XPATH, "//div[@id='quickClassOptions']/a").click()
        time.sleep(2)
        
        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH,"//div[@id='tabsDiv']/ul/li[4]/a"))).click()
        
        WebDriverWait(browser, 100).until(EC.presence_of_all_elements_located ((By.XPATH, "(//table[@class='GMSection']/tbody)[4]/tr[@class='GMDataRow']")))
        
        try:                        
            if len(browser.find_elements(By.XPATH, "//tr[@class='GMDataRow']/td[6]"))>0:
                mpns=browser.find_elements(By.XPATH, "//tr[@class='GMDataRow']/td[5]/a[@class='image_link']")
                time.sleep(2)
                for mpn in mpns:
                    time.sleep(1)
                    mpn.click()                    
                    time.sleep(2)
                    WebDriverWait(browser, 30).until(EC.element_to_be_clickable((By.XPATH,"//div[@id='tabsDiv']/ul/li[4]/a"))).click()
                    WebDriverWait(browser, 30).until(EC.element_to_be_clickable((By.XPATH,"//a[@id='MSG_AddAttachment_2']"))).click()
                    time.sleep(3)
                    inputfile = browser.find_element(By.CSS_SELECTOR,"#add-files-input")    
                    browser.execute_script("arguments[0].style.visibility = 'visible'; arguments[0].style.height = '1px'; arguments[0].style.width = '1px'; arguments[0].style.opacity = 1", inputfile)           
                    # change_visibility = '$("#add-files-input").css("visibility,"visible");'
                    # change_display = '$("#add-files-input").css("display,"block");'
                    # browser.execute_script(change_visibility)
                    # browser.execute_script(change_display)
                    time.sleep(3)
                    
                    time.sleep(3)
                    for fileS in os.listdir(download_dir):
                        filename = os.path.join(download_dir,fileS)
                        time.sleep(5)
                        inputfile.send_keys(filename)
                        time.sleep(5)
                        browser.find_element(By.XPATH,"//a[@id='uploadFilesUM']").click()
                        time.sleep(10)
        except:
            print("No print")
        time.sleep(2)
        i=i+1

    except:
        print("break")
