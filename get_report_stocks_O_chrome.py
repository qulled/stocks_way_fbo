import datetime as dt
import os
import pickle
import time

from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium_stealth import stealth
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


options = Options()

prefs = {'download.default_directory': r'excel_docs/'}

options.add_experimental_option('prefs', prefs)
options.add_argument("--disable-blink-features")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
options.add_argument('--headless')

driver = webdriver.Chrome(options=options,service=Service(ChromeDriverManager().install()))

driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
driver.execute_cdp_cmd('Network.setUserAgentOverride', {
    "userAgent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.5060.53 Safari/537.36'})

stealth(driver,
        languages=["ru-Ru", "en"],
        vendor="Google Inc.",
        platform="Win32",
        webgl_vendor="Intel Inc.",
        renderer="Intel Iris OpenGL Engine",
        fix_hairline=True,
        )


BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def auth(url,name):
    driver.get(url)
    cookies = pickle.load(open(f'cookies-{name}.py', 'rb'))
    for cookie in cookies:
        driver.add_cookie(cookie)
    time.sleep(20)
    attempt = driver.find_element(By.CLASS_NAME,'WarningCookiesBannerCard__button__DSLFl2gcQr')
    attempt.click()
    return time.sleep(1)



def get_report(name):
    setting_table_button = driver.find_element(By.CLASS_NAME,'Warehouse-remains__button-setting__1lRrtaruNg')
    setting_table_button.click()
    time.sleep(2)
    article_button = driver.find_element(By.XPATH,'/html/body/div[4]/div[3]/div/div/div[3]/div[3]/div/label/span')
    article_button.click()
    time.sleep(0.5)
    nomenclarure_button = driver.find_element(By.XPATH, '/html/body/div[4]/div[3]/div/div/div[3]/div[4]/div/label/span')
    nomenclarure_button.click()
    time.sleep(0.5)
    barcode_button = driver.find_element(By.XPATH, '/html/body/div[4]/div[3]/div/div/div[3]/div[5]/div/label/span')
    barcode_button.click()
    time.sleep(0.5)
    size_button = driver.find_element(By.XPATH, '/html/body/div[4]/div[3]/div/div/div[3]/div[6]/div/label/span')
    size_button.click()
    time.sleep(0.5)
    save_setting_button = driver.find_element(By.XPATH,'/html/body/div[4]/div[3]/div/div/div[3]/button')
    save_setting_button.click()
    time.sleep(5)
    download_button = driver.find_element(By.CLASS_NAME,'Warehouse-remains__button-excel__1kzRdUbKae')
    download_button.click()
    time.sleep(7)
    if int(day) < 10:
        file_oldname = os.path.join(dirparth, f'report_{year}_{month.strip("0")}_{day.strip("0")}.xlsx')
        file_newname = os.path.join(dirparth, f'{name}-{year}-{month}-{day}.xlsx')
        os.rename(file_oldname, file_newname)
    else:
        file_oldname = os.path.join(dirparth, f'report_{year}_{month.strip("0")}_{day}.xlsx')
        file_newname = os.path.join(dirparth, f'{name}-{year}-{month}-{day}.xlsx')
        os.rename(file_oldname, file_newname)
    return time.sleep(3)


if __name__ == '__main__':
    day = dt.datetime.now().strftime('%d')
    month = dt.datetime.now().strftime("%m")
    year = dt.datetime.now().strftime("%Y")
    dirparth = r'excel_docs/'
    try:
        name = 'Орлова'
        auth('https://seller.wildberries.ru/analytics/warehouse-remains',name)
        get_report(name)
    except Exception as e:
        print(e)
    finally:
        driver.quit()
        exit()
