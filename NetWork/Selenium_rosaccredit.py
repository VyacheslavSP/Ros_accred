from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
import keyboard
import pyautogui
import win32clipboard
import win32api
import win32com.client


def main_insert_accredit(excell_path, only_write, only_send_tmp):
    options = Options()
    options.add_argument("--test-type")
    options.add_argument(
        '--user-data-dir=C:/Users/VecheslavSP/AppData/Local/Google/Chrome/User Data/Default')
    options.add_argument('--allow-plugins')
    options.add_argument('--profile-directory=Profile 1')
    driver = webdriver.Chrome(options=options)
    driver.get('http://10.250.74.17/auth/')
    button = driver.find_element(By.CLASS_NAME, 'login')
    button.click()
    time.sleep(1)
    try:
        try:
            driver.switch_to.window(driver.window_handles[1])
            time.sleep(0.3)
            ph = driver.find_element(
                By.ID, 'login')
            ph.send_keys('+79654199169')
        except:
            print("без логина")
        try:
            pas = driver.find_element(
                By.ID, 'password')
            pas.send_keys('T9z$)pbK8~@')
            pas.send_keys(Keys.ENTER)
            time.sleep(0.3)
        except:
            print("без пароля")
        but_later = driver.find_element(By.CLASS_NAME, 'plain-button-inline')
        but_later.click()
    except:
        print("без госуслуг")
    driver.switch_to.window(driver.window_handles[1])
    time.sleep(0.3)
    rostest = driver.find_element(By.TAG_NAME, 'td')
    rostest.click()
    time.sleep(0.3)
    while (True):
        driver.get('http://10.250.74.17/lk/activities')
        time.sleep(0.3)
        try:
            driver.find_element(By.CLASS_NAME, 'access-denied__text')
            time.sleep(0.5)
            driver.get('http://10.250.74.17/lk')
            time.sleep(0.5)
        except:
            time.sleep(0.5)
            break
    while (True):
        try:
            art = driver.find_element(By.TAG_NAME, 'article')
            art.click()
            time.sleep(0.2)
            break
        except:
            time.sleep(0.5)
    while (True):
        try:
            poverk = driver.find_element(
                By.XPATH, '/html/body/fgis-root/div/fgis-roei/fgis-roei-attestation/fgis-roei-reestr-selector/div/div[4]')
            poverk.click()
            break
        except:
            time.sleep(0.5)
    while (True):
        try:
            elem = (driver.find_elements(By.TAG_NAME, 'tbody'))
            if (len(elem) != 0):
                break
            else:
                raise
        except:
            time.sleep(1)
    body = driver.find_element(By.NAME, 'model')
    body.send_keys('')
    if (only_write == True):
        try_whith_selenium_action(driver, excell_path)
    if (only_send_tmp == True):
        acsept_value(excell_path)
    if (only_write and only_send_tmp):
        try_whith_selenium_action(driver, excell_path)
        acsept_value(excell_path)
    else:  # ветка нового варианта
        button = driver.find_element(
            By.CSS_SELECTOR, "body > fgis-root > div > fgis-roei > fgis-roei-verification-measuring-instruments > div > div > div.header-block > fgis-table-toolbar > section > div > div.left-side > div > fgis-toolbar > div > div:nth-child(5) > fgis-toolbar-button > button")
        button.click()
        entry_body = driver.find_element(By.ID, 'file')
        entry_body.send_keys(
            'C:/Users/VecheslavSP/Desktop/Python/Ros_accred/ex_corr/export.xml')
        time.sleep(0.5)  # запас на тупление сайта
        button_2 = driver.find_element(
            By.CSS_SELECTOR, '#mainDialog > fgis-modal > div > div.fgis-modal__content > div.fgis-modal__footer > div > button:nth-child(1)')
        button_2.click()
        return driver


def read_txt_write_console():
    with open("for_rosaccredit.txt", "r", encoding="UTF-8", errors='replace') as file:
        str1 = file.read()
    print(str1)
    win32api.LoadKeyboardLayout('00000419', 1)
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(str1)
    win32clipboard.CloseClipboard()
    win32api.LoadKeyboardLayout('00000409', 1)
    return


def change_log(sheet, i):
    f = open('C:/Users/VecheslavSP/Desktop/Python/Ros_accred/ex_corr/log_number_send.txt',
             'r', encoding='utf-8')
    value_array = []
    while (True):
        line = f.readline()
        if not line:
            break
        value_array.append(line.strip())
    nameTostr = time.localtime()
    name = str(nameTostr[2])+"."+str(nameTostr[1])+"."+str(nameTostr[0]
                                                           )+"."+str(nameTostr[3])+";"+str(nameTostr[4])
    data = str(sheet.Cells(
        i+2, 1).value).split()[0]+"*"+name+"*"+str(sheet.Cells(i+2, 6).value).split()[0]
    value_array.append(data)
    f.close
    f = open('C:/Users/VecheslavSP/Desktop/Python/Ros_accred/ex_corr/log_number_send.txt',
             'w', encoding='utf-8')
    for index in range(len(value_array)):
        f.write(str(value_array[index])+'\n')
    f.close


def try_whith_console():
    keyboard.send("F12")
    time.sleep(2)
    pyautogui.click(541, 208)
    pyautogui.click(541, 208)
    read_txt_write_console()
    keyboard.send("Ctrl+V")
    keyboard.send("Enter")


def get_maximum_rows(sheet):
    rows = 0
    i = 1
    while i < 10000:
        if (sheet.Cells(i, 1).value != None):
            rows += 1
            i += 1
        else:
            break
    return rows

# <div _ngcontent-ejt-c5="" class="waiter" hidden=""><div _ngcontent-ejt-c5="" class="waiter__content"><div _ngcontent-ejt-c5="" class="waiter__image"></div><div _ngcontent-ejt-c5="" class="waiter__text">Загрузка...</div></div></div>


def check_wait(driver):
    try:
        flag = False
        while (True):  # делай пока не надоест
            for j in range(5):  # 5 попыток
                try:
                    loader = driver.find_element(By.CLASS_NAME, "waiter")
                    flag = True
                except:
                    if (flag):
                        flag = not (flag)
                        break
            return
    except:
        time.sleep(2)
        print("Ошибка поиска загрузгиW")
        return


def acsept_value(path_excel):  # отправить все черновики
    driver = main_insert_accredit(path_excel, False, False)
    Excel = win32com.client.Dispatch("Excel.Application")
    wb = Excel.Workbooks.Open(path_excel)
    sheet = wb.ActiveSheet
    rows = get_maximum_rows(sheet)
    for i in range(rows-1):
        check_wait(driver)
        pole = driver.find_element(By.CSS_SELECTOR, "body > fgis-root > div > fgis-roei > fgis-roei-verification-measuring-instruments > div > fgis-roei-verification-measuring-instruments-advanced-search > fgis-filters-panel > fgis-left-panel > div.left-panel_body > div.body > div > div.filter__item > fgis-field-input > fgis-field-wrapper > div > div > input")
        pole.clear()
        time.sleep(2)
        pole.send_keys(sheet.Cells(i+2, 1).value)
        time.sleep(2)
        driver.find_element(By.CSS_SELECTOR, "body > fgis-root > div > fgis-roei > fgis-roei-verification-measuring-instruments > div > fgis-roei-verification-measuring-instruments-advanced-search > fgis-filters-panel > fgis-left-panel > div.left-panel_footer > div > button").click()
        time.sleep(5)
        check_wait(driver)
        driver.find_element(By.CSS_SELECTOR, "body > fgis-root > div > fgis-roei > fgis-roei-verification-measuring-instruments > div > div > div.container-table > fgis-table > div.header-wrapper > div > div.table-head-fex.div-tr > div.table-head-fex.table-check.div-th.div-th_first.ng-star-inserted > p-checkbox > div > div.ui-chkbox-box.ui-widget.ui-corner-all.ui-state-default").click()
        time.sleep(2)
        driver.find_element(By.CSS_SELECTOR, "body > fgis-root > div > fgis-roei > fgis-roei-verification-measuring-instruments > div > div > div.header-block > fgis-table-toolbar > section > div > div.left-side > div > fgis-toolbar > div > div:nth-child(4) > fgis-toolbar-button > button").click()
        time.sleep(2)
        driver.find_element(
            By.CSS_SELECTOR, "body > fgis-root > div > fgis-roei > fgis-roei-verification-measuring-instruments > fgis-modal > div > div.fgis-modal__content > div.fgis-modal__footer > div > button").click()
        time.sleep(10)
        check_wait(driver)
        change_log(sheet, i)
    wb.Save()
    time.sleep(0.5)
    wb.Close()
    time.sleep(0.5)
    Excel.Quit()


def try_whith_selenium_action(driver, excell_path):
    Excel = win32com.client.Dispatch("Excel.Application")
    wb = Excel.Workbooks.Open(excell_path)
    sheet = wb.ActiveSheet
    rows = get_maximum_rows(sheet)
    good = True
    for i in range(rows-1):  # убрать заголовок
        time.sleep(1)
        button = driver.find_element(
            By.XPATH, '/html/body/fgis-root/div/fgis-roei/fgis-roei-verification-measuring-instruments/div/div/div[1]/fgis-table-toolbar/section/div/div[1]/div/fgis-toolbar/div/div[1]/fgis-toolbar-button/button')
        button.click()
        time.sleep(0.5)
        driver.find_element(By.CSS_SELECTOR, 'body > fgis-root > div > fgis-roei > fgis-verification-measuring-instruments-card-edit > div > div > div > div > fgis-verification-measuring-instruments-card-edit-common > fgis-card-block > div > div.card-block__container > div > fgis-card-edit-row-two-columns:nth-child(1) > fgis-card-edit-row:nth-child(1) > div.card-edit-row__content > fgis-field-input > fgis-field-wrapper > div > div > input').send_keys(sheet.Cells(i+2, 1).value)
        driver.find_element(By.CSS_SELECTOR, 'body > fgis-root > div > fgis-roei > fgis-verification-measuring-instruments-card-edit > div > div > div > div > fgis-verification-measuring-instruments-card-edit-common > fgis-card-block > div > div.card-block__container > div > fgis-card-edit-row-two-columns:nth-child(1) > fgis-card-edit-row:nth-child(2) > div.card-edit-row__content > fgis-field-input > fgis-field-wrapper > div > div > input').send_keys(sheet.Cells(i+2, 2).value)
        driver.find_element(By.CSS_SELECTOR, 'body > fgis-root > div > fgis-roei > fgis-verification-measuring-instruments-card-edit > div > div > div > div > fgis-verification-measuring-instruments-card-edit-common > fgis-card-block > div > div.card-block__container > div > fgis-card-edit-row:nth-child(3) > div.card-edit-row__content > fgis-field-selectbox > fgis-field-wrapper > div > div > fgis-selectbox > div').send_keys(Keys.ENTER)
        time.sleep(1.5)
        if str(sheet.Cells(i+2, 5).value).split()[0] == "Пригодно":
            good = True
            driver.find_element(
                By.CSS_SELECTOR, 'body > fgis-root > fgis-select-dropdown > div > div > div.fgis-selectbox__filter.ng-star-inserted > input').send_keys(Keys.ENTER)
        else:
            good = False
            time.sleep(2)
            driver.find_element(
                By.CSS_SELECTOR, 'body > fgis-root > fgis-select-dropdown > div > div > div.fgis-selectbox__filter.ng-star-inserted > input').send_keys(Keys.DOWN)
            time.sleep(2)
            driver.find_element(
                By.CSS_SELECTOR, 'body > fgis-root > fgis-select-dropdown > div > div > div.fgis-selectbox__filter.ng-star-inserted > input').send_keys(Keys.ENTER)
        driver.find_element(By.CSS_SELECTOR, 'body > fgis-root > div > fgis-roei > fgis-verification-measuring-instruments-card-edit > div > div > div > div > fgis-verification-measuring-instruments-card-edit-common > fgis-card-block > div > div.card-block__container > div > fgis-card-edit-row:nth-child(5) > div.card-edit-row__content > fgis-field-selectbox > fgis-field-wrapper > div > div > fgis-selectbox > div').send_keys(Keys.ENTER)
        time.sleep(2)
        driver.find_element(
            By.CSS_SELECTOR, 'body > fgis-root > fgis-select-dropdown > div > div > div.fgis-selectbox__filter.ng-star-inserted > input').send_keys(sheet.Cells(i+2, 6).value)
        time.sleep(4)
        # вставить блок для однофамилцев. Жесто зашить проверку на имя???

        driver.find_element(By.CSS_SELECTOR, 'body > fgis-root > div > fgis-roei > fgis-verification-measuring-instruments-card-edit > div > div > div > div > fgis-verification-measuring-instruments-card-edit-common > fgis-card-block > div > div.card-block__container > div > fgis-card-edit-row:nth-child(5) > div.card-edit-row__content > fgis-field-selectbox > fgis-field-wrapper > div > div > fgis-selectbox > div').send_keys(Keys.ENTER)
        time.sleep(2)
        driver.find_element(By.CSS_SELECTOR, 'body > fgis-root > div > fgis-roei > fgis-verification-measuring-instruments-card-edit > div > div > div > div > fgis-verification-measuring-instruments-card-edit-common > fgis-card-block > div > div.card-block__container > div > fgis-card-edit-row-two-columns:nth-child(2) > fgis-card-edit-row:nth-child(1) > div.card-edit-row__content > fgis-field-calendar > fgis-field-wrapper > div > div > fgis-calendar > div > div > input').send_keys(sheet.Cells(i+2, 3).value)
        time.sleep(5)
        if (good):  # если пригодно то 2 дата
            driver.find_element(By.CSS_SELECTOR, 'body > fgis-root > div > fgis-roei > fgis-verification-measuring-instruments-card-edit > div > div > div > div > fgis-verification-measuring-instruments-card-edit-common > fgis-card-block > div > div.card-block__container > div > fgis-card-edit-row-two-columns:nth-child(2) > fgis-card-edit-row:nth-child(2) > div.card-edit-row__content > fgis-field-calendar > fgis-field-wrapper > div > div > fgis-calendar > div > div > input').send_keys(sheet.Cells(i+2, 4).value)
            time.sleep(2)
        driver.find_element(
            By.CSS_SELECTOR, 'body > fgis-root > div > fgis-roei > fgis-verification-measuring-instruments-card-edit > fgis-verification-measuring-instruments-card-edit-toolbar > div > fgis-toolbar > div > div:nth-child(1) > fgis-toolbar-button > button').click()
        time.sleep(3)
        while (True):
            try:
                button = driver.find_element(
                    By.XPATH, '//*[@id="mainDialog"]/fgis-modal/div/div[1]/div[3]/div/button')
                button.click()
                break
            except:
                time.sleep(2)
    wb.Save()
    time.sleep(0.5)
    wb.Close()
    time.sleep(0.5)
    Excel.Quit()
    driver.close()  # закрыть обе вкладки
    driver.close()


# main_insert_accredit("123", False, False)
