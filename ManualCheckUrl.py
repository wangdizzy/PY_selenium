from selenium.webdriver.chrome.service import Service as ChromeService
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from skpy.chat import SkypeChats
from skpy import Skype
import gspread
import requests
import gspread


"""
檢查sta sta2
"""
print('-----start-----')
message = ""

scopes = [
    "https://spreadsheets.google.com/feeds"
]  # 定義存取的範圍 feeds = google sheet

credentials = ServiceAccountCredentials.from_json_keyfile_name(
    "sheet.json", scopes
)  # 指定檔案金鑰

client = gspread.authorize(credentials)  # 傳入gspread模組

sheet = client.open_by_key(
    "1qWdc0QTGY13LEsr_N_5SA4cfyWTHaGXk8GhJFosoNzc"
).sheet1  # 使用open_by_key方式傳入google sheet 金鑰


print("環境：1.STA 2.STA2")
num = input("輸入1 or 2：")
url = list()
urlrange = [
    "D",
    "E",
    "F",
    "G",
    "H",
    "I",
    "J",
    "K",
    "L",
    "M",
    "N",
    "O",
    "P",
    "Q",
    "R",
    "S",
    "T",
    "U",
    "V",
    "W",
    "X",
    "Y",
    "Z",
    "AA",
    "AB",
    "AC",
    "AD",
]
while num not in "12" and len(num) > 1 and num is None:
    print("輸入內容非1 or 2")
    num = input("請重新輸入數字：")
if num == "1":
    # 讀取google sheet內容
    testaccount = sheet.acell("A3").value
    pswd = sheet.acell("B3").value
    env = sheet.acell("C3").value  # 從google sheet取得欄位C3資料
    for x in urlrange:
        thisurl = sheet.acell("%s3" % x).value
        if thisurl is not None:
            url.append(thisurl)
        else:
            break
else:
    # 讀取google sheet內容
    testaccount = sheet.acell("A4").value
    pswd = sheet.acell("B4").value
    env = sheet.acell("C4").value  # 從google sheet取得欄位C4資料
    for x in urlrange:
        thisurl = sheet.acell("%s4" % x).value
        if thisurl is not None:
            url.append(thisurl)
        else:
            break

options = Options()  # 設定Options
options.add_argument(
    "user-agent= Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.5112.79 Safari/537.36"
)  # 指定user-agent
options.add_experimental_option("excludeSwitches", ["enable-logging"])
options.add_argument("--start-maximized")  # 視窗最大
chrome = webdriver.Chrome(
    service=ChromeService(ChromeDriverManager(driver_version="126.0.6478.183").install()), options=options
)

# 用request打URL 確認status code
for x in url:
    try:
        statuscode = requests.get(x)
    except:
        print("請確認網址是否輸入正確 %s" % x)
        continue
    if statuscode.status_code == 200:
        chrome.get(x)
        if "newsmart.368aa" in x or "newsmart.cmmd368" in x:  # sta 新手機
            try:
                chrome.implicitly_wait(5)  # 隱含等待
                chrome.find_element(By.NAME, "name").send_keys(testaccount)
                chrome.find_element(By.NAME, "pwd").send_keys(pswd)
                chrome.find_element(By.ID, "btn-login").click()
                chrome.implicitly_wait(5)  # 隱含等待
                # 判斷是否要更新密碼
                oldPwd = chrome.find_element(By.XPATH, '//*[@id="view-home"]/div[2]/div/div/div[1]').text
                if oldPwd != '':
                    message += "密碼需要更新%s\n" % x
                else:
                    chrome.implicitly_wait(5)  # 隱含等待
                    chrome.find_element(
                        By.XPATH, '//*[@id="link-sport"]/div[1]/img').click()
                    chrome.find_element(
                        By.XPATH, '//*[@id="toolbar-sport"]/div/a[4]/i').click()
                    chrome.find_element(
                        By.XPATH,"/html/body/div[7]/div[1]/div/ul/li[3]/a/div[2]/div[1]",).click()
                    print("%s 可以正常切換Market Type(新手機)" % x)
                    continue
            except:
                print("%s 有異常請確認" % x)
                message += "發生異常，前往下列網址確認服務是否正常 %s\n" % x
                continue
        elif "smart.368aa" in x or "smart.cmmd368" in x:  # sta 舊手機
            try:
                chrome.implicitly_wait(5)
                chrome.find_element(By.NAME, "username").send_keys(testaccount)
                chrome.find_element(By.NAME, "password").send_keys(pswd)
                chrome.find_element(By.ID, "btnLogin").click()
                chrome.implicitly_wait(5)  # 隱含等待
                # 判斷是否要更新密碼
                try:
                    oldPwd = chrome.find_element(By.XPATH, '//*[@id="action"]').text
                    message += "密碼需要更新%s\n" % x
                except:
                    chrome.find_element(By.XPATH, '//*[@id="asportpanelmenu"]').click()
                    print("%s 可以正常登入(舊手機)" % x)
                    continue
            except:
                print("%s 有異常請確認" % x)
                message += "發生異常，前往下列網址確認服務是否正常 %s\n" % x
                continue
        else:  # web
            try:
                if "max222" in x:
                    chrome.implicitly_wait(5)
                    chrome.find_element(By.ID, "txtID").send_keys(testaccount)
                    chrome.find_element(By.ID, "txtPW").send_keys(pswd)
                    chrome.find_element(By.ID, "login").click()
                elif "gc855" in x:
                    chrome.implicitly_wait(5)
                    chrome.find_element(By.ID, "UserName").send_keys(testaccount)
                    chrome.find_element(By.ID, "Password").send_keys(pswd)
                    chrome.find_element(By.ID, "btnLogin").click()
                else:
                    chrome.implicitly_wait(5)
                    chrome.find_element(By.ID, "UserName").send_keys(testaccount)
                    chrome.find_element(By.ID, "Password").send_keys(pswd)
                    chrome.find_element(By.ID, "sub").click()
                chrome.implicitly_wait(5)  # 隱含等待
                try:  # 判斷是否要更新密碼
                    oldPwd = chrome.find_element(
                        By.XPATH, '//*[@id="tr_oldpwd"]/th'
                    ).text
                    message += "密碼需要更新%s\n" % x
                except:
                    try:
                        chrome.find_element(
                            By.XPATH, '//*[@id="divTradeInAd"]/div/div/div/p/button'
                        ).click()
                    except:
                        chrome.implicitly_wait(5)  # 隱含等待
                        try:  # 判斷Betview
                            if bool(chrome.find_element(By.XPATH, "/html/body")):
                                pass
                            else:
                                message += "無法取得BetView %s\n" % x
                        except:
                            message += "無法取得BetView %s\n" % x 
                    try:
                        chrome.implicitly_wait(5)  # 隱含等待
                        chrome.switch_to.frame("leftFrame")  # 切換frame
                        chrome.find_element(By.ID, "btn_staus").click()
                        balance = chrome.find_element(By.ID, "lb_balance").text.replace(",", "")  # 取得Balance金額
                        outstanding = chrome.find_element(
                            By.ID, "lb_outstanding").text.replace(",", "")  # 取得Outstanding金額
                        print("正常取得用戶金額 %s" % x)
                    except:
                        message += "無法取得用戶金額 %s\n" % x
            except:
                if chrome.title == "System Maintenance":
                    message += "%s is System Maintenance." % x
                else:
                    print("發生異常 %s" % EOFError)
                    message += "發生異常，前往下列網址確認服務是否正常 %s\n" % x

    elif statuscode.status_code == 403:
        print("請確認VPN是否連線")
    else:
        message += "發生異常，前往下列網址確認服務是否正常 %s\n" % x

if len(message) > 0:
    # 登入帳號
    skUser = sheet.acell("A8").value
    skPw = sheet.acell("B8").value
    sk = Skype(user=skUser, pwd=skPw)

    # 建立SK物件
    skc = SkypeChats(sk)

    # 指定發送群組
    cht_name = sheet.acell("C8").value

    # 群組ID
    room_id = None

    # 獲取前10個群組資料
    chats = skc.recent()

    while room_id is None and len(chats) > 0:
        # 使用遍歷將資料取出
        for x in chats.values():
            # 取得群組名稱和id
            group_name = getattr(x, "topic", "no attr")
            group_id = getattr(x, "id", "no id")

            # 當名稱相同時執行
            if group_name in cht_name:
                # 將group_id給到room_id後結束迴圈
                room_id = group_id
                break
        chats = skc.recent()

    ch = sk.chats[room_id]
    sendMessage = ch.sendMsg("%s 站台連線檢查結果如下\n%s" % (env, message))
print("-----end-----")
chrome.quit()