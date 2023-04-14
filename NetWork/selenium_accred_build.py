from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time


def build_accredit_console(path_to_excel):
    options = Options()
    link = "http://app.arshin-fsa.ru:8080/"
    driver = webdriver.Chrome(options=options)
    driver.get(link)
    driver.find_element(By.XPATH, "//*[@id='excelRadio']")
    Radio = driver.find_element(By.ID, 'excelRadio')
    Radio.click()
    time.sleep(0.5)
    driver.implicitly_wait(10)
    path_ex = driver.find_element(
        By.XPATH, '/html/body/div/div/form/div/div[3]/div[2]/div/input')
    path_ex.send_keys(path_to_excel)
    button_click = driver.find_element(
        By.XPATH, '/html/body/div/div/form/div/div[4]/button')
    button_click.click()
    time.sleep(0.5)
    return


# https://arshin-fsa.ru/
