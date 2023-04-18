from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
import win32clipboard
import win32api
import win32com.client


def get_maximum_rows(sheet):
    rows = 0
    i = 1
    while i < 10000:
        if (sheet.Cells(i, 1).Value != None):
            rows += 1
            i += 1
        else:
            break
    return rows


def find_reestr(path_of_excell):
    options = Options()
    # options.add_argument("--headless=new")
    Excel = win32com.client.Dispatch("Excel.Application")
    wb = Excel.Workbooks.Open(path_of_excell)
    sheet = wb.ActiveSheet
    rows = get_maximum_rows(sheet)
    for i in range(rows-1):        # -1 строка заголовка

        link = "https://fgis.gost.ru/fundmetrology/cm/results?filter_result_docnum=" + \
            sheet.Cells(i+2, 1).Value.split()[0]

        driver = webdriver.Chrome(options=options)
        driver.get(link)
        elem = []
        flag = True
        while (flag):
            if (str(driver.get_log('browser')[0]['message']).find("status of 429", 0, -1) != -1):
                time.sleep(0.2)
                driver.refresh()
            else:
                while (True):
                    try:
                        elem = (driver.find_elements(By.TAG_NAME, 'td'))
                        time.sleep(0.2)
                        if (len(elem) != 0):
                            sheet.Cells(i+2, 2).Value = elem[1].text
                            sheet.Cells(i+2, 3).Value = elem[6].text
                            try:
                                # если парситтся по точке
                                if (elem[7].text != None):
                                    sheet.Cells(i+2, 4).Value = elem[7].text
                            except:
                                None
                            driver.quit()
                            flag = False
                            break
                        else:
                            str_tmp = driver.find_elements(
                                By.CLASS_NAME, 'row')[11].text
                            if (str_tmp.find("По Вашему запросу", 0, len(str_tmp)) != -1):
                                driver.refresh()
                    except:
                        time.sleep(0.2)
    wb.Save()
    wb.Close()
    Excel.Quit()
    return


def get_valid_number():
    numbers_from_list = []
    original_number = []
    f = open('C:\\Users\\VecheslavSP\\Desktop\\Python\\Ros_accred\\ex_corr\\number_send.txt',
             'r', encoding='utf-8')
    for line in f:
        if (line != ''):
            original_number.append(line)
            line = str(line).split('/')
            numbers_from_list.append(line[2])
    return numbers_from_list, original_number


def get_valid_value():
    array_numbers = []
    numbers_from_list = get_valid_number()[0]
    for elem in numbers_from_list:
        index = numbers_from_list.index(elem)
        array_numbers.append(find_reestr(numbers_from_list[index]))
    return array_numbers


def send_valid_value(numbers_from_list, array_numbers):
    f = open('C:\\Users\\VecheslavSP\\Desktop\\Python\\Ros_accred\\ex_corr\\valid_value_send.txt',
             'w', encoding='utf-8')
    for index in range(len(numbers_from_list)):
        f.write(str(numbers_from_list[index]).strip() +
                '*'+str(array_numbers[index])+'\n')
    return


def Main_res(path_of_excell):
    find_reestr(path_of_excell)
    return


# Main_res()
