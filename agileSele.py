import sys
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
from selenium.webdriver.common.action_chains import ActionChains
from dotenv import load_dotenv

EDGE_DRIVER = r'msedgedriver.exe'



# env_path='.env'
#load_dotenv(dotenv_path=env_path)
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

browser.get(r'https://iam-signin.juniper.net/app/juniper_agileplm_1/exk1mv21lsoJQuzDl0h8/sso/saml')

# main_page = browser.current_window_handle

wait = WebDriverWait(browser, 300)
waitLoginURL = r"https://iam-signin.juniper.net"

wait.until(EC.url_contains(waitLoginURL))
nameWait = WebDriverWait(browser, 20)

nameWait.until(EC.visibility_of_any_elements_located((By.ID, r"idp-discovery-username")))
browser.find_element(By.NAME, r"username").send_keys()
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
browser.find_element(By.XPATH, r"//input[@type='submit' and @value='Sign in']").click()

time.sleep(3)
# https://agileplm.juniper.net/Agile/default/login-cms.jsp
browser.get(r'https://agileplm.juniper.net/Agile/PLMServlet')

browser.switch_to.window(browser.window_handles[1])
wait3 = WebDriverWait(browser, 300)
wait3LoginURL = r"https://agileplm.juniper.net/Agile/PLMServlet"

wait3.until(EC.url_contains(wait3LoginURL))

browser.get(r'https://agileplm.juniper.net/Agile/PLMServlet')

# collapse LeftPane
# browser.find_element(By.XPATH, "//div[@id='collapse']").click()
# time.sleep(2)


# Search
browser.find_element(By.XPATH, "//input[@id='QUICKSEARCH_STRING']").send_keys("310-167452"+ Keys.ENTER)
time.sleep(3)

## JPN Edit 
editButton = browser.find_element(By.XPATH, "//a[@id='MSG_Edit']")
browser.execute_script("arguments[0].click()", editButton)
time.sleep(2)
# # browser.find_element(By.XPATH, "(//a[@class='button'])[12]").click()
# # time.sleep(1)

# JPN edit Title Page 
desCription = browser.find_element(By.XPATH ,"//textarea[@id='R1_1002_0']")
desCription.send_keys("Item Description to fill")
time.sleep(2)

Select(browser.find_element(By.ID, "R1_1082_0")).select_by_visible_text("Buy".strip())
time.sleep(2)
browser.find_element(By.XPATH ,"//input[@id='search_query_R1_1004_0_display']").send_keys('SRX-forge', Keys.ENTER)
time.sleep(2)

Select(browser.find_element(By.ID, "R1_66318_11")).select_by_visible_text("Omkar Chogle".strip())
Select(browser.find_element(By.ID, "R1_2000018808_11")).select_by_visible_text("Debayan Dutta".strip())
Select(browser.find_element(By.ID, "R1_2000018806_11")).select_by_visible_text("Hyman Pei".strip())
Select(browser.find_element(By.ID, "R1_88262_11")).select_by_visible_text("Multisourcing Rule Not Applicable".strip())
Select(browser.find_element(By.ID, "R1_66328_11")).select_by_visible_text("Single Source".strip())
Select(browser.find_element(By.ID, "R1_66321_11")).select_by_visible_text("No".strip())
Select(browser.find_element(By.ID, "R1_2027_11")).select_by_visible_text("Yes".strip())
time.sleep(2)

browser.find_element(By.XPATH ,"//input[@id='R1_2000018766_11']").send_keys('5', Keys.TAB)
browser.find_element(By.XPATH ,"//input[@id='R1_2000018767_11']").send_keys('9', Keys.TAB)
browser.find_element(By.XPATH ,"//input[@id='R1_2000018768_11']").send_keys('9', Keys.TAB)
browser.find_element(By.XPATH ,"//input[@id='R1_87713_11']").send_keys('0.5', Keys.TAB)
time.sleep(5)

saveButton = browser.find_element(By.XPATH, "//a[@id='MSG_Save']")
browser.execute_script("arguments[0].click()", saveButton)
time.sleep(6)

# browser.find_element(By.XPATH ,"//div[@id='tabsDiv']/ul/li[2]/a").click()
time.sleep(6)
# AVL Tab
browser.find_element(By.XPATH ,"//div[@id='tabsDiv']/ul/li[4]/a").click()
time.sleep(4)

addButton = browser.find_element(By.XPATH ,"//a[@id='MSG_Add_9']")
browser.execute_script("arguments[0].click()", addButton)


browser.find_element(By.XPATH ,"//a[@id='create_C3']").click()
time.sleep(1)

browser.switch_to.window(browser.window_handles[2])
time.sleep(2)
mfgName = browser.find_element(By.CSS_SELECTOR, "input#search_query_manufacturername_display")
mfgName.send_keys("TEXAS INSTRUMENTS",Keys.ENTER)
mfgNumber = browser.find_element(By.XPATH, "//input[@id='number']")
mfgNumber.send_keys("TIDMO6543", Keys.ENTER)
riskReason = browser.find_element(By.XPATH, "//div[@id='edit_mode_R1_2000008073_4_display']/div/input[@id='search_query_R1_2000008073_4_display']")
riskReason.send_keys("0", Keys.TAB)
browser.find_element(By.XPATH ,"//div[@class='column_three']/a[@id='add']").click()
browser.switch_to.window(browser.window_handles[1])
time.sleep(2)

## if doesn't exist
browser.find_element(By.XPATH, "//input[@id='QUICKSEARCH_STRING']").clear()
browser.find_element(By.XPATH, "//input[@id='QUICKSEARCH_STRING']").send_keys("TIDMO6543"+ Keys.ENTER)
editMPN = browser.find_element(By.XPATH, "//div[@class ='rt_column']/p/a[@class='button']")
browser.execute_script("arguments[0].click()", editMPN)

val = browser.find_element(By.XPATH, "//ul[@id='selected_items_R1_2091_4_display']")
if val.text == '':
    roHSButton = browser.find_element(By.XPATH, "//div[@id ='edit_mode_R1_2091_4_display']/div/input[@id='search_query_R1_2091_4_display']")
    roHSButton.send_keys("No E", Keys.TAB)
if val.text == 'Not Determined':
    browser.find_element(By.XPATH, "//a[@id='_R1_2091_4_display_1_close']").click()
else:
    roHSButton = browser.find_element(By.XPATH, "//div[@id ='edit_mode_R1_2091_4_display']/div/input[@id='search_query_R1_2091_4_display']")
    roHSButton.send_keys("No E", Keys.TAB)

browser.find_element(By.XPATH, "//dd[@id='col_1301']/input[@id='R1_1301_4']").send_keys("85", Keys.TAB)
browser.find_element(By.XPATH, "//dd[@id='col_1302']/input[@id='R1_1302_4']").send_keys("100", Keys.TAB)
browser.find_element(By.XPATH, "//dd[@id='col_1303']/input[@id='R1_1303_4']").send_keys("125", Keys.TAB)
Select(browser.find_element(By.XPATH, "//dd[@id='col_1272']/select[@id='R1_1272_4']")).select_by_visible_text("Yes".strip())
browser.find_element(By.XPATH, "//input[@id='search_query_R1_2090_4_display']").send_keys("Matte", Keys.TAB)
browser.find_element(By.XPATH, "//dd[@id='col_2090']/div/div/input[@id='search_query_R1_2090_4_display']").send_keys("Mat", Keys.TAB)
Select(browser.find_element(By.XPATH, "//dd[@id='col_1275']/select[@id='R1_1275_4']")).select_by_visible_text("Fully Compliant".strip())
browser.find_element(By.XPATH, "//div[@id ='edit_mode_R1_2091_4_display']/div/input[@id='search_query_R1_2091_4_display']")
MPNsavButton = browser.find_element(By.XPATH, "//a[@id='MSG_Save']")
browser.execute_script("arguments[0].click()", MPNsavButton)

## JPN atachments
browser.find_element(By.XPATH ,"//div[@id='tabsDiv']/ul/li[7]/a").click()

browser.close()
browser.quit()