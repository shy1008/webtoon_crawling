from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import openpyxl
import time

from operator import length_hint

from webdriver_manager.chrome import ChromeDriverManager

# HTTPS
PROXY = "xxx.xx.xxx.xxx:xxxx" #ip:port


webdriver.DesiredCapabilities.CHROME['proxy'] = {
    "httpProxy": PROXY,
    "ftpProxy": PROXY,
    "sslProxy": PROXY,
    "proxyType": "MANUAL"
}
webdriver.DesiredCapabilities.CHROME['acceptSslCerts']=True

chrome_options = Options()
chrome_options.add_experimental_option("detach", True) #브라우저 꺼짐 방지 
chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"]) #불필요한 에러메세지 없앰


wb = openpyxl.Workbook()
sheet = wb.active
sheet.append(["site_addr", "site_title","con"])

service = Service(executable_path=ChromeDriverManager().install())
browser = webdriver.Chrome(service=service,options=chrome_options)

browser.get('htpps://xxx') # url
browser.implicitly_wait(30)

browser.execute_script('window.scrollTo(0, document.body.scrollHeight);')
browser.implicitly_wait(10)

items = browser.find_elements(By.CSS_SELECTOR, ".tile.col-md-6")

for item in items:
    time.sleep(1)

    title = item.find_element(By.CSS_SELECTOR, ".desc > h3").text
    # print(title)
    link = item.find_element(By.CSS_SELECTOR, ".desc > h3 > a").get_attribute('href')
    # print(link)

    browser.execute_script(f"window.open('{link}');")
    time.sleep(3)

    browser.switch_to.window(browser.window_handles[-1])
    time.sleep(1)

    try:
        counts = browser.find_elements(By.XPATH , "//*[@id='chapters-list']/table/tbody/tr")
        count = len(counts)
        # print(count)
    except:
        print("error except")
        count = 999

    sheet.append([link,title,count])
    time.sleep(1)

    browser.switch_to.window(browser.window_handles[0])
    time.sleep(1)


wb.save("crawling.xlsx")
print("종료")
