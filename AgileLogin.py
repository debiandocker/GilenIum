from json import JSONDecodeError
import pandas as pd
import os
import time
import requests
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from msedge.selenium_tools import Edge, EdgeOptions
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
from selenium.common.exceptions import NoSuchElementException
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
browser.find_element(By.XPATH, r"//input[@type='submit' and @value='Sign in']").click()

time.sleep(3)
# https://agileplm.juniper.net/Agile/default/login-cms.jsp
browser.get(r'https://agileplm.juniper.net/Agile/PLMServlet')

browser.switch_to.window(browser.window_handles[1])
wait3 = WebDriverWait(browser, 300)
wait3LoginURL = r"https://agileplm.juniper.net/Agile/PLMServlet"

wait3.until(EC.url_contains(wait3LoginURL))

browser.get(r'https://agileplm.juniper.net/Agile/PLMServlet')

wait3.until(EC.visibility_of_any_elements_located((By.LINK_TEXT, r"Create New")))
button=browser.find_element(By.CSS_SELECTOR, "a#toggle_create_menu")
browser.execute_script("arguments[0].click();",button)
time.sleep(2)
# wait3.until(lambda driver: driver.execute_script(r"return jQuery.active == 0"))
partlink = browser.find_element(By.XPATH, r"//div[@id='901']/div/ul/li/a[@href]")
browser.execute_script(r"arguments[0].click();",partlink)
time.sleep(2)

# Select dropdown Category
browser.switch_to.window(browser.window_handles[2])
browser.set_window_size(700,600)
time.sleep(10)
select = Select(browser.find_element(By.ID, r"subClassId"))
select.select_by_visible_text(r"J-310 Semiconductors".strip())
time.sleep(2)

# Generate Autonumer 
try:
    partNumber = browser.find_element(By.CSS_SELECTOR, "a.button")
    browser.execute_script("arguments[0].click();",partNumber)
except KeyError:
    pass
time.sleep(2)

# Prefix Commodity code
eleme = browser.find_element(By.CSS_SELECTOR, "#R1_1001_0").get_attribute('value').replace('XX0',"310")
browser.find_element(By.CSS_SELECTOR, "#R1_1001_0").clear()
browser.find_element(By.CSS_SELECTOR, "#R1_1001_0").send_keys(eleme)

browser.find_element(By.CSS_SELECTOR, "#search_query_R1_1282_11_display").send_keys("31031",Keys.TAB)

# Part Creation Saved
button2=browser.find_element(By.XPATH, "//a[@id='save']")
browser.execute_script("arguments[0].click();",button2)
time.sleep(20)


# changed Window
browser.switch_to.window(browser.window_handles[1])
time.sleep(3)

# Save Part
button3=browser.find_element(By.XPATH, "//*[@id='MSG_Save']")
browser.execute_script("arguments[0].click();",button3)



time.sleep(10)


browser.close()
browser.quit()