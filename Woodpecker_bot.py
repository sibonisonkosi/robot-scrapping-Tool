from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import pickle
import json
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
# package for sharepoint
from shareplum import Office365
from shareplum import Site
from shareplum.site import Version

from requests.auth import HTTPBasicAuth
import Logic

# to be used variables
failed_campaigns = []
campaigns_with_steps = []
count = 0

# Change this when you about to host to Luna
# PATH = 'C:\Program Files (x86)\chromedriver.exe'
# driver = webdriver.Chrome(PATH)
driver = webdriver.Chrome(ChromeDriverManager().install())
# start login code
url_login = 'https://app.woodpecker.co/panel?u=2090#campaigns/1221943/stats'  # used as a login URL
# to be used stats url with a string holder "%s"
url_campaign_edit = 'https://app.woodpecker.co/panel?u=2090#campaigns/%s/stats'

NpEncoder = Logic.NpEncoder()
obj_ = Logic.Methods(driver, url_login, url_campaign_edit)

obj_.launch_login_page()
obj_.create_session() # Retrieve cookies if you have entered the credentials once, if not there will be no session to re-call
time.sleep(10)
# ============= Emails ===========
obj_.extracting_email_response()
driver.quit()


print("===================Done===============================================")

