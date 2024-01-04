from selenium.webdriver.support.ui import Select
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from skpy.chat import SkypeChats
from skpy import Skype
import gspread
import requests
import time
import gspread



 #登入帳號
sk = Skype(user='kevin.wang@collaborate.tw', pwd='Dizzy9951')

#建立SK物件
skc = SkypeChats(sk)

#指定發送群組
cht_name = '批次訊息'

#群組ID
room_id = None

#獲取前10個群組資料
chats = skc.recent()

while room_id is None and len(chats) > 0:
    #使用遍歷將資料取出
    for x in chats.values():
        #取得群組名稱和id
        group_name = getattr(x, 'topic', 'no attr')
        group_id = getattr(x, 'id', 'no id')
        
        #當名稱相同時執行
        if group_name in cht_name:
            #將group_id給到room_id後結束迴圈
            room_id = group_id
            break
    chats = skc.recent()
    
ch = sk.chats[room_id]
message = ''

scopes = ['https://spreadsheets.google.com/feeds'] #定義存取的範圍 feeds = google sheet

credentials = ServiceAccountCredentials.from_json_keyfile_name('sheet.json', scopes) #指定檔案金鑰

client = gspread.authorize(credentials) #傳入gspread模組

sheet = client.open_by_key('1qWdc0QTGY13LEsr_N_5SA4cfyWTHaGXk8GhJFosoNzc').sheet1 #使用open_by_key方式傳入google sheet 金鑰



url = list()
urlrange = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD']
for num in range(1,3):
    if num == 1: #thor
        #讀取google sheet內容
        testaccount = sheet.acell('A2').value 
        pswd = sheet.acell('B2').value
        env = sheet.acell('C2').value #從google sheet取得欄位C2資料
        for x in urlrange: 
            thisurl = sheet.acell('%s2' %x).value 
            if thisurl is not None:
                url.append(thisurl) 
            else:
                break
    else: #sta
        #讀取google sheet內容
        testaccount = sheet.acell('A3').value 
        pswd = sheet.acell('B3').value
        env = sheet.acell('C3').value #從google sheet取得欄位C3資料
        for x in urlrange:
            thisurl = sheet.acell('%s3' %x).value 
            if thisurl is not None:
                url.append(thisurl)
            else:
                break
            
    options = Options() #設定Options
    options.add_argument('--no-sandbox')
    options.add_argument(
        'user-agent= Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.5112.79 Safari/537.36') #指定user-agent
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_argument('--start-maximized') #視窗最大
    chrome = webdriver.Chrome(ChromeDriverManager().install(), options=options)

    #用request打URL 確認status code
    for x in url:
        try:
            statuscode = requests.get(x)
        except:
            message +=  '請確認網址是否輸入正確 %s' % x
            continue
        if statuscode.status_code == 200:
            chrome.get(x)
            time.sleep(1)
            if 'newsmart.368aa' in x: #sta 新手機
                try:
                    chrome.find_element_by_name('name').send_keys(testaccount)
                    chrome.find_element_by_name('pwd').send_keys(pswd)    
                    chrome.find_element_by_id('btn-login').click()
                    time.sleep(3)
                    chrome.find_element_by_xpath('//*[@id="link-sport"]/div[1]/img').click()
                    time.sleep(1)
                    chrome.find_element_by_xpath('//*[@id="toolbar-sport"]/div/a[4]/i').click()
                    time.sleep(1)
                    chrome.find_element_by_xpath('/html/body/div[7]/div[1]/div/ul/li[3]/a/div[2]/div[1]').click()
                    print('%s 可以正常切換Market Type(新手機)' % x)
                    continue
                except:
                    print('%s 有異常請確認' % x)
                    message += '發生異常，前往下列網址確認服務是否正常 %s\n' % x
                    continue
                
            elif 'smart.12vin' in x: #thor 新手機
                try:
                    chrome.find_element_by_name('name').send_keys(testaccount)
                    chrome.find_element_by_name('pwd').send_keys(pswd)    
                    chrome.find_element_by_id('btn-login').click()
                    time.sleep(3)
                    chrome.find_element_by_xpath('//*[@id="link-sport"]/div[1]/img').click()
                    time.sleep(1)
                    chrome.find_element_by_xpath('//*[@id="toolbar-sport"]/div/a[4]/i').click()
                    time.sleep(1)
                    chrome.find_element_by_xpath('/html/body/div[7]/div[1]/div/ul/li[3]/a/div[2]/div[1]').click()
                    print('%s 可以正常切換Market Type(新手機)' % x)
                    continue
                except:
                    print('%s 有異常請確認' % x)
                    message += '發生異常，前往下列網址確認服務是否正常 %s\n' % x
                    continue
               
            #elif 'mobile.cmdbet' in x: #prod 新手機
            #    try:
            #        chrome.find_element_by_name('name').send_keys(testaccount)
            #        chrome.find_element_by_name('pwd').send_keys(pswd)    
            #        chrome.find_element_by_id('btn-login').click()
            #        time.sleep(3)
            #        chrome.find_element_by_xpath('//*[@id="link-sport"]/div[1]/img').click()
            #        time.sleep(1)
            #        chrome.find_element_by_xpath('//*[@id="toolbar-sport"]/div/a[4]/i').click()
            #        time.sleep(1)
            #        chrome.find_element_by_xpath('/html/body/div[7]/div[1]/div/ul/li[3]/a/div[2]/div[1]').click()
            #        print('%s 可以正常切換Market Type(新手機)' % x)
            #        continue
            #    except:
            #        print('%s 有異常請確認' % x)
            #        message += '發生異常，前往下列網址確認服務是否正常 %s\n' % x
            #        continue
                 
            elif 'mobile.12vin' in x: #thor 舊手機
                try:
                    chrome.find_element_by_name('username').send_keys(testaccount)
                    chrome.find_element_by_name('password').send_keys(pswd)   
                    chrome.find_element_by_id('btnLogin').click()
                    time.sleep(3)
                    chrome.find_element_by_xpath('//*[@id="asportpanelmenu"]').click()
                    time.sleep(3)
                    chrome.find_element_by_xpath('//*[@id="betTypeList"]/li[7]/a').click()
                    print('%s 可以正常切換Market Type(舊手機)' % x)
                    continue
                except:
                    print('%s 有異常請確認' % x)
                    message += '發生異常，前往下列網址確認服務是否正常 %s\n' % x
                    continue
                
            elif 'smart.368aa' in x: #sta 舊手機
                try:
                    chrome.find_element_by_name('username').send_keys(testaccount)
                    chrome.find_element_by_name('password').send_keys(pswd)   
                    chrome.find_element_by_id('btnLogin').click()
                    time.sleep(3)
                    chrome.find_element_by_xpath('//*[@id="asportpanelmenu"]').click()
                    time.sleep(3)
                    chrome.find_element_by_xpath('//*[@id="betTypeList"]/li[7]/a').click()
                    print('%s 可以正常切換Market Type(舊手機)' % x)
                    continue
                except:
                    print('%s 有異常請確認' % x)
                    message += '發生異常，前往下列網址確認服務是否正常 %s\n' % x
                    continue
                
            #elif 'smart.cmdbet' in x: #prod 舊手機
            #    try:
            #        chrome.find_element_by_name('username').send_keys(testaccount)
            #        chrome.find_element_by_name('password').send_keys(pswd)   
            #        chrome.find_element_by_id('btnLogin').click()
            #        time.sleep(3)
            #        chrome.find_element_by_xpath('//*[@id="asportpanelmenu"]').click()
            #        time.sleep(3)
            #        chrome.find_element_by_xpath('//*[@id="betTypeList"]/li[7]/a').click()
            #        print('%s 可以正常切換Market Type(舊手機)' % x)
            #        continue
            #    except:
            #        print('%s 有異常請確認' % x)
            #        message += '發生異常，前往下列網址確認服務是否正常 %s\n' % x
            #        continue
                
            else:    #web

                #selectLangue = Select(chrome.find_element_by_id('ddl_language'))
                #selectLangue.select_by_value('en-US')
                #time.sleep(1)
                try:
                    if 'MAXBET' in chrome.title:
                        chrome.find_element_by_id('txtID').send_keys(testaccount)
                        chrome.find_element_by_id('txtPW').send_keys(pswd)
                        time.sleep(1)
                        chrome.find_element_by_id('login').click()
                    else:
                        chrome.find_element_by_id('UserName').send_keys(testaccount)
                        chrome.find_element_by_id('Password').send_keys(pswd)
                        time.sleep(1)
                        chrome.find_element_by_id('sub').click()
                    time.sleep(1)
                    try:
                        chrome.find_element_by_xpath('//*[@id="divTradeInAd"]/div/div/div/p/button').click()
                    except:
                        pass    
                    chrome.switch_to.frame('leftFrame') #切換frame
                    chrome.find_element_by_id('btn_staus').click()
                    time.sleep(1)
                    try:
                        balance = chrome.find_element_by_id('lb_balance').text.replace(',','') #取得Balance金額
                        outstanding = chrome.find_element_by_id('lb_outstanding').text.replace(',','') #取得Outstanding金額
                        betcredit = chrome.find_element_by_id('lb_bet_credit').text.replace(',','') #取得Bet Crtedit金額
                        print('正常取得用戶金額 %s' %x)
                    except:
                        message += '無法取得用戶金額 %s' % x
                except:
                    if chrome.title == "System Maintenance":
                        message += '%s is System Maintenance.' % x
                    else:
                        message += '發生異常，前往下列網址確認服務是否正常 %s\n' % x

        elif statuscode.status_code == 403:
            print('請確認VPN是否連線')
        else:
            message += '發生異常，前往下列網址確認服務是否正常 %s\n' % x

    chrome.close()
    url = []
    if len(message) > 0:
        sendMessage = ch.sendMsg('%s 站台連線檢查結果如下\n%s' % (env, message))