from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options  
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import time, threading, schedule, pickle, os

# 登入資訊
USERNAME = "wxbireader@chainwin-agrotech.com"
PASSWORD = "zM9WEsS3"    

# Power BI 網址
URLS = [
    "https://app.powerbi.com/groups/f7739ba9-bfd5-499c-8dd5-79768925ff2b/reports/2d3afc19-9fc0-4f1d-b23b-f3959e6ea29b/3ea62ee9b9406bea99ce?experience=power-bi",
    "https://app.powerbi.com/groups/f7739ba9-bfd5-499c-8dd5-79768925ff2b/reports/2d3afc19-9fc0-4f1d-b23b-f3959e6ea29b/41a1b1dc0ee719a4bb52?experience=power-bi",
    "https://app.powerbi.com/groups/f7739ba9-bfd5-499c-8dd5-79768925ff2b/reports/2d3afc19-9fc0-4f1d-b23b-f3959e6ea29b/a29114746ceca0982e4e?experience=power-bi"
]

COOKIE_PATH = "cookies.pkl"
drivers = []

# 等待報表畫面完成載入
def wait_until_report_ready(driver, timeout=30):
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.reportContent, div.visual-container"))
        )
        print("🟢 報表已載入完成")
    except:
        print("🔴 報表載入超時")

# 全螢幕（先點擊報表畫面，確保焦點正確）
def enter_fullscreen(driver):
    try:
        body = driver.find_element(By.TAG_NAME, "body")
        body.click()
        time.sleep(0.5)
        actions = ActionChains(driver)
        actions.key_down(Keys.CONTROL).key_down(Keys.SHIFT).send_keys('f').key_up(Keys.SHIFT).key_up(Keys.CONTROL).perform()
        print("🟢 已進入全螢幕")
    except Exception as e:
        print(f"⚠️ 全螢幕失敗：{e}")

# 建立 selenium driver
def get_driver():
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-infobars") 
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])  
    chrome_options.add_experimental_option('useAutomationExtension', False)

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=chrome_options
    )
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver

# 儲存 / 載入 cookie
def save_cookies(driver):
    with open(COOKIE_PATH, "wb") as f:
        pickle.dump(driver.get_cookies(), f)

def load_cookies(driver):
    with open(COOKIE_PATH, "rb") as f:
        cookies = pickle.load(f)
        for cookie in cookies:
            if 'sameSite' in cookie and cookie['sameSite'] == 'None':
                cookie['sameSite'] = 'Strict'
            driver.add_cookie(cookie)

# 自動登入（若 cookie 無效）
def login_microsoft(driver):
    try:
        driver.get("https://login.microsoftonline.com/")
        time.sleep(3)
        driver.find_element(By.NAME, "loginfmt").send_keys(USERNAME + Keys.ENTER)
        time.sleep(3)
        driver.find_element(By.NAME, "passwd").send_keys(PASSWORD + Keys.ENTER)
        time.sleep(3)
        try:
            stay_signed_in = driver.find_element(By.XPATH, "//input[@type='submit' and contains(@id, 'idBtn')]")
            stay_signed_in.click()
            print("保持登入狀態")
        except:
            print("略過保持登入")
    except Exception as e:
        print(f"⚠️ 登入錯誤: {e}")

# 驗證登入是否成功
def is_login_success(driver):
    time.sleep(5)
    current_url = driver.current_url
    return "powerbi.com" in current_url and "report" in current_url

# 開啟所有 Power BI 頁面
def open_pages():
    for url in URLS:
        driver = get_driver()
        driver.get("https://app.powerbi.com/")
        time.sleep(3)

        login_success = False

        if os.path.exists(COOKIE_PATH):
            print("🟡 嘗試載入 cookie")
            try:
                load_cookies(driver)
                driver.get(url)
                if is_login_success(driver):
                    print("✅ 使用 cookie 成功登入")
                    login_success = True
                else:
                    print("⚠️ Cookie 已失效，將重新登入")
            except Exception as e:
                print(f"❌ 載入 cookie 錯誤：{e}")

        if not login_success:
            login_microsoft(driver)
            driver.get(url)
            if is_login_success(driver):
                print("✅ 重新登入成功，儲存新 cookie")
                save_cookies(driver)
            else:
                print("❌ 登入失敗，略過此頁")
                driver.quit()
                continue

        wait_until_report_ready(driver)
        time.sleep(2)
        enter_fullscreen(driver)
        drivers.append(driver)
        print(f"{datetime.now()} ✅ 開啟並全螢幕：{url}")

# 定時重新整理
def refresh_pages():
    for driver in drivers:
        try:
            driver.refresh()
            print(f"{datetime.now()} 🔁 已重新整理")
            wait_until_report_ready(driver)
            time.sleep(2)
            enter_fullscreen(driver)
        except Exception as e:
            print(f"{datetime.now()} ⚠️ 重新整理失敗：{e}")

# 排程執行
def schedule_loop():
    refresh_times = ["00:00", "04:00", "08:00", "12:00", "16:00", "20:00"]
    for t in refresh_times:
        schedule.every().day.at(t).do(refresh_pages)
    print("🕒 已設定排程：", ", ".join(refresh_times))

    while True:
        schedule.run_pending()
        time.sleep(10)

# 主程序
if __name__ == "__main__":
    open_pages()
    schedule_loop()
