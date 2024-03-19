from selenium.webdriver.support.ui import Select
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
import gspread
import time

'''
規則

'''


scopes = ['https://spreadsheets.google.com/feeds'] #定義存取的範圍 feeds = google sheet

credentials = ServiceAccountCredentials.from_json_keyfile_name('sheet.json', scopes) #指定檔案金鑰

client = gspread.authorize(credentials) #傳入gspread模組

sheet = client.open_by_key('1A3Mb2Jz_JnrZUFE5zCeGS4kdF11beTqeH-i-4Pi49lY') #使用open_by_key方式傳入google sheet 金鑰
worksheet = sheet.get_worksheet(0)

options = Options()
options.add_argument('--no-sandbox')
options.add_argument(
    'user-agent= Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.5112.79 Safari/537.36') #指定user-agent
options.add_experimental_option('excludeSwitches', ['enable-logging'])
options.add_argument('--start-maximized') #視窗最大
chrome = webdriver.Chrome(ChromeDriverManager().install(), options=options)

chrome.get('https://cmdmember.12vin.com/')

chrome.find_element_by_id('UserName').send_keys('l700011222333')
chrome.find_element_by_id('Password').send_keys('1qaz@WSX')
chrome.find_element_by_id('sub').click()

chrome.switch_to_frame('topFrame')
chrome.find_element_by_id('help').click()

# 跳出 Frame
chrome.switch_to.default_content()  

# 移動到How To Bet頁面
chrome.switch_to.window(chrome.window_handles[1])  

#進入 static-topframe
chrome.switch_to_frame('static-topframe')

#點how to bet
chrome.find_element_by_xpath('/html/body/div/div[1]/div/a[2]').click() 

chrome.switch_to.default_content()  

#進入 leftFrame
chrome.switch_to_frame('leftFrame') 

try:
    sport = worksheet.acell('C2').value
    chrome.find_element_by_link_text(sport).click()
    chrome.switch_to.default_content()  
    chrome.switch_to_frame('mainFrame') 
except:
    print('沒有找到%s' % sport)

for x in range(1,100):
    jira = sheet.get_worksheet(1).acell('A%s' % x).value
    print(jira)
    chrome.find_element_by_xpath("//span[contains(text(),'%s')]" % jira)


chrome.quit()