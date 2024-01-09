'''
海大自動搶課系統
'''
## 延遲時間相關
import base64
import io
import logging
import os
import time
import traceback
## 單元測試模組，線性測試用不到
import unittest
import warnings

import ddddocr
import matplotlib.image as mpimg
import pytesseract
## BeautifulReport: 產生自動測試報告套件
from BeautifulReport import BeautifulReport
from matplotlib import pyplot as plt
from PIL import Image, ImageEnhance, ImageFilter
from selenium import webdriver
from selenium.common.exceptions import *
## ActionChains: 滑鼠事件相關
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.alert import *
## BY: 也就是依照條件尋找元素中XPATH、CLASS NAME、ID、CSS選擇器等都會用到的Library
from selenium.webdriver.common.by import By
## keys: 鍵盤相關的Library
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.remote_connection import LOGGER
## expected_conditions: 條件相關
from selenium.webdriver.support import expected_conditions as EC
## WebDriverWait: 等待頁面加載完成的顯性等待機制Library
## Select: 下拉選單相關支援，但前端框架UI工具不適用(ex: Quasar、ElementUI、Bootstrap)
from selenium.webdriver.support.ui import Select, WebDriverWait
## Chrome WebDriver 需要DRIVER Manager的支援
from webdriver_manager.chrome import ChromeDriverManager

## 設定Chrome的瀏覽器彈出時遵照的規則
## 這串設定是防止瀏覽器上頭顯示「Chrome正受自動控制」
options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
## 關閉自動記住密碼的提示彈窗
options.add_experimental_option("prefs", {
    "profile.password_manager_enabled": False,
    "credentials_enable_service": False
})
#不開啟實際網頁
# options.add_argument('--headless')

#關閉訊息
options.add_experimental_option('excludeSwitches', ['enable-logging'])
options.add_argument("--log-level=0")
options.add_argument('--disable-extensions')
options.add_argument('test-type')

FORMAT = '%(asctime)s %(filename)s %(levelname)s:%(message)s'
logging.basicConfig(level=logging.DEBUG,
                    format=FORMAT,
                    filename='Log.log',
                    filemode='a')


#驗證碼破解物件
class verifyCodeBreaker:
    #驗證碼破解(有機率會出錯)
    def hack(self, driver):

        #抓取網頁驗證碼圖片
        img_base64 = driver.execute_script(
            """
            var ele = arguments[0];
            var cnv = document.createElement('canvas');
            cnv.width = ele.width; cnv.height = ele.height;
            cnv.getContext('2d').drawImage(ele, 0, 0);
            return cnv.toDataURL('image/jpeg').substring(22);
            """, driver.find_element(By.XPATH, "//img[@id='importantImg']"))
        #儲存圖片
        with open("captcha_login.png", 'wb') as image:
            image.write(base64.b64decode(img_base64))

        #解析驗證碼(海大驗證碼一定都是數字+大寫英文字母)
        ocr = ddddocr.DdddOcr()
        with open('captcha_login.png', 'rb') as f:
            img_bytes = f.read()
        #圖片識別驗證碼
        res = ocr.classification(img_bytes)
        #刪掉圖片
        os.remove("captcha_login.png")

        return res.upper()


#使用者資料
class Information:

    def __init__(self):
        self.__username = input("請輸入帳號: ")
        self.__password = input("請輸入密碼: ")

    #獲得使用者帳號
    def get_username(self):
        return self.__username

    #獲得使用者密碼
    def get_password(self):
        return self.__password


#網頁驅動class
class WebDriver:
    #constructor
    def __init__(self):
        self.driver = webdriver.Chrome(chrome_options=options)
        #定義操作串鍊變數
        self.action = ActionChains(self.driver)
        #爬蟲目標網頁
        self.URL = "https://ais.ntou.edu.tw/Default.aspx"
        self.driver.get(self.URL)

        #等待載入頁面完畢
        while (True):
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located(
                        (By.CLASS_NAME, "form-label")))
                break
            except TimeoutError:
                print("載入頁面失敗，正在重新嘗試...")
            except Exception as err:
                logging.warning("載入網頁出現問題 => " + str(err))
                print("網頁無回應，請等待網頁正常再開啟....")
                return

        self.user = None
        self.driver.maximize_window()

    #輸入登入資訊
    def Login(self):

        #建立使用者資料
        self.user = Information()
        #抓取帳號輸入框
        username_textbox = self.driver.find_element(By.ID,
                                                    "M_PORTAL_LOGIN_ACNT")
        #清除帳號框內文字
        username_textbox.clear()

        #抓取密碼輸入框
        password_textbox = self.driver.find_element(By.ID, "M_PW")

        #抓取驗證碼輸入框
        verifyCode_textbox = self.driver.find_element(By.ID, "M_PW2")

        #抓取登入按鈕
        login_btn = self.driver.find_element(By.ID, "LGOIN_BTN")

        #輸入資訊
        username_textbox.send_keys(self.user.get_username())
        password_textbox.send_keys(self.user.get_password())
        verifyCode_textbox.send_keys(verifyCodeBreaker().hack(self.driver))

        #按下登入按鈕
        login_btn.click()

        #檢查是否登入錯誤
        flag = True
        try:
            alert = self.driver.switch_to.alert
            alert.accept()
            print("請檢查帳號密碼是否輸入正確")
            flag = False
        except Exception as err:
            print("---------------------登入成功---------------------")

        return flag

    #切換至即時加退選頁面
    def change(self):
        #等待頁面載入
        try:
            WebDriverWait(self.driver, 3).until(
                EC.presence_of_element_located((By.ID, "menuFrame")))
            #切換至側邊欄框架
            sidebar_frame = self.driver.find_element(
                By.XPATH, "//iframe[@id='menuIFrame']")
            self.driver.switch_to.frame(sidebar_frame)
        except TimeoutError:
            print("載入頁面失敗")
            return

        #點擊教務系統
        try:
            WebDriverWait(self.driver, 3).until(
                EC.presence_of_element_located((By.ID, "Menu_TreeViewt1")))
            self.driver.find_element(By.ID, "Menu_TreeViewt1").click()
        except TimeoutError:
            print("點擊教務系統按鈕無回應")
            return

        #點擊選課系統
        try:
            WebDriverWait(self.driver, 3).until(
                EC.presence_of_element_located((By.ID, "Menu_TreeViewt31")))
            self.driver.find_element(By.ID, "Menu_TreeViewt31").click()
        except TimeoutError:
            print("點擊選課系統按鈕無回應")
            return

        #點擊線上即時加退選
        try:
            WebDriverWait(self.driver, 3).until(
                EC.presence_of_element_located((By.ID, "Menu_TreeViewt41")))
            self.driver.find_element(By.ID, "Menu_TreeViewt41").click()
        except TimeoutError:
            print("點擊線上即時加退選按鈕無回應")
            return

        self.driver.switch_to.default_content()

        #切換至線上即時加退選主頁面
        try:
            WebDriverWait(self.driver, 3).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//iframe[@id='mainIFrame']")))
            instantly_select_course_main_frame = self.driver.find_element(
                By.XPATH, "//iframe[@id='mainIFrame']")
            self.driver.switch_to.frame(instantly_select_course_main_frame)
        except TimeoutError:
            pass

    #加選課程
    def selecting(self, course):

        #課號搜尋文字框
        try:
            WebDriverWait(self.driver, 3).until(
                EC.presence_of_element_located((By.ID, "Q_COSID")))
            #抓取課號搜尋文字框物件
            course_id_textbox = self.driver.find_element(By.ID, "Q_COSID")
        except TimeoutError:
            pass
        #搜尋課程
        course_id_textbox.clear()
        course_id_textbox.send_keys(course[0])
        #送出課號搜尋按鈕
        course_id_btn = self.driver.find_element(By.ID, "QUERY_COSID_BTN")
        course_id_btn.click()

        #搜尋結果table並等待結果出現
        try:
            search_result_table = self.driver.find_element(
                By.ID, "DataGrid1").find_elements(By.TAG_NAME, "tr")
            WebDriverWait(self.driver,
                          3).until(EC.staleness_of(search_result_table[1]))
        except TimeoutError:
            pass
        #重新抓取最新的table
        search_result_table = self.driver.find_element(By.ID, "DataGrid1")

        #目標加選按鈕
        result_btn = None
        #紀錄搜尋結果第幾列
        result_idx = 0
        #找到目標欄位
        for row in search_result_table.find_elements(By.TAG_NAME, "tr")[1:]:
            cols = row.find_elements(By.TAG_NAME, "td")
            result_idx += 1
            if (cols[2].text == course[0] and cols[3].text == course[1]):
                result_btn = cols[0].find_element(By.TAG_NAME, "a")
                break

        comp = (course[0] + ' ' + course[1])

        print("正在嘗試加選 " + comp + ".....")

        #不斷嘗試加入課程直到加選成功
        while (True):

            #按下加選按鈕
            try:
                webdriver.ActionChains(self.driver).move_to_element(
                    result_btn).click(result_btn).perform()
                self.resolveAllAlerts(2, True)
                if (self.driver.find_element(By.ID,
                                             "Div2").text.__contains__(comp)):
                    break
            except StaleElementReferenceException:
                #重新抓取最新的table
                search_result_table = self.driver.find_element(
                    By.ID, "DataGrid1")
                #目標加選按鈕
                result_btn = search_result_table.find_elements(
                    By.TAG_NAME, "tr")[result_idx].find_elements(
                        By.TAG_NAME, "td")[0].find_element(By.TAG_NAME, "a")
            except TimeoutException:
                print("網頁無回應，請等待網頁正常再執行......")
                return
            except UnexpectedAlertPresentException:
                self.resolveAllAlerts(2, True)

        print(comp + " 加選成功!")
        #B92B1G04 B

    #檢查是否有alert
    def isAlertPresent(self, timeout):
        try:
            #等待alert跳出
            alert = WebDriverWait(self.driver,
                                  timeout).until(EC.alert_is_present())

            if (not alert):
                raise TimeoutException
            return True
        except TimeoutException:
            return False

    #處理多次的alert
    def resolveAllAlerts(self, timeout, accept):
        while (self.isAlertPresent(timeout)):
            self.resolveAlert(accept)
            time.sleep(1)

    #分別處理alert
    def resolveAlert(self, accept):

        if (accept):
            self.driver.switch_to.alert.accept()
        else:
            self.driver.switch_to.alert.dismiss()


if __name__ == '__main__':
    print(
        "---------------------歡迎使用海大搶課系統，請登入您的海大教學務系統之帳號密碼---------------------"
    )
    crawler = WebDriver()

    #登入
    while (not crawler.Login()):
        pass
    print("*注意:請確認欲加選的課程不衝堂或出現其他無法加選的情況..............程式會報錯")

    print("請輸入欲加選的課號以及班別(ex:B57030TX A)，並輸入end退出")

    targets = []
    while (True):
        cmd = input()
        if cmd == "end":
            break

        targets.append(cmd)
    print("---------------------輸入完畢---------------------")

    #切換頁面至即時加退選頁面
    crawler.change()

    print("---------------------正在加選課程中---------------------")

    for target in targets:
        crawler.selecting(target.split(' '))

    print("---------------------課程均加選完畢---------------------")
    print()
    print("正在退出系統......")
