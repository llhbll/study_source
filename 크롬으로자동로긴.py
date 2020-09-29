# 999 포탈 자동로그인 IE는 에러가 계속나서 전자결재는 OPEN안함
import os
import sys
import time

from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

#pyinstaller --noconsole --onefile --add-binary "chromedriver.exe";"." gmail_auto.py
if getattr(sys, 'frozen', False):
    chromedriver_path = os.path.join(sys._MEIPASS, "chromedriver.exe")
    driver = webdriver.Chrome(chromedriver_path)
else:
    driver = webdriver.Chrome()

url = 'https://portal.sungshin.ac.kr/sso/login.jsp'
driver.get(url)
action = ActionChains(driver)
driver.implicitly_wait(10)
driver.find_element_by_css_selector('#loginId_mobile').click()

action.send_keys('2970021').key_down(Keys.TAB).send_keys('lchlshp12*').key_down(Keys.TAB).key_down(Keys.TAB).key_down(Keys.ENTER).perform()
driver.implicitly_wait(10)
time.sleep(1)
driver.get('https://tis.sungshin.ac.kr/comm/nxui/staff/sso.do?menuUid=PORTAL_3201&connectDiv=1')
driver.maximize_window()