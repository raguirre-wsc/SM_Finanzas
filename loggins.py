from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import tkinter.simpledialog as simpledialog
from pywinauto import *
from pywinauto.keyboard import send_keys
import pyautogui as rpa
import arrow





def log_IB():
    path= r"C:\Users\rodriaguirre\Desktop\chromedriver_win32\chromedriver2"

    global driver
    driver=webdriver.Chrome(path)

    driver.get("https://sib1.interbanking.com.ar/secureLogin.do?from=home")

    search=driver.find_element_by_name("documento")
    search.send_keys("20-39549930-1")

    search.send_keys(Keys.RETURN)

    def grab_id(id):
        try:
            search = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, id))
            )
            return search
        except:
            driver.quit()


    grab_id("username").send_keys("rawsc123")
    grab_id("password").send_keys("inter1234")

    grab_id("password").send_keys(Keys.RETURN)

    cod = simpledialog.askstring("Data input window","Token:")

    grab_id("token").send_keys(cod)
    grab_id("token").send_keys(Keys.RETURN)

    return driver

def out():
    path = r"C:\Users\rodriaguirre\Desktop\chromedriver_win32\chromedriver2"

    global driver
    driver = webdriver.Chrome(path)

    driver.get("https://sib1.interbanking.com.ar/secureLogin.do?from=home")

    def grab_id(id):
        try:
            search = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, id))
            )
            return search
        except:
            print("not found")
            driver.quit()

    while True:
        grab_id('/html/body/table/tbody/tr[2]/td[3]/img').click()
        print("clickie")
        time.sleep(2)

    return driver

def zcl04():
    SAP=Application().start(r"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe")
    time.sleep(7)
    for x in range(7):
        send_keys("{DOWN}")
    send_keys("{ENTER}")
    time.sleep(7)
    send_keys("zcl04")
    send_keys("{ENTER}")
    time.sleep(2)
    send_keys("+{F5}")
    time.sleep(2)
    send_keys("pagosposidiari")
    for x in range(4):
        send_keys("{TAB}")
        time.sleep(0.05)
    for x in range(15):
        send_keys("{DELETE}")
    send_keys("{F8}")
    for x in range(28):
        send_keys("{TAB}")
        time.sleep(0.05)
    tadd2 = arrow.now().shift(days=2).format("DD.MM.YYYY")
    tsus5 = arrow.now().shift(days=-5).format("DD.MM.YYYY")
    send_keys(tsus5)
    send_keys("{TAB}")
    send_keys(tadd2)
    for x in range(35):
        send_keys("{TAB}")
    send_keys("{DELETE}")
    send_keys("{F8}")




