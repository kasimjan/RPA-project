from pyPythonRPA.Robot import pythonRPA
from selenium import webdriver
import os
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from time import sleep
from selenium.webdriver.chrome.options import Options
import pyautogui
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from Sources.other_functions import check_ie_driver_settings
import keyboard

check_ie_driver_settings()

caps = DesiredCapabilities.INTERNETEXPLORER
caps['ignoreProtectedModeSettings'] = True
caps['InternetExplorerDriver.INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS'] = True

url = "http://bpmd/"
username = "mazhit.e"
password = "Asd_12345678"

driver = webdriver.Ie(r"C:\Windows\System32\IEDriverServer.exe", capabilities=caps)

#names = ['Абдраханова Гульмира Дилжановна', 'Абдраимов Канат Турсынбаевич', 'Абдраймов Рахимжан Амирбекович']

def __selenium_wait_element(by, selector, sec = 2, appear=True):
    """Ожидание элемента"""
    sleep(0.3)
    try:
        if appear:
            WebDriverWait(driver, sec).until(EC.presence_of_element_located((by, selector)))
        else:
            WebDriverWait(driver, sec).until_not(EC.presence_of_element_located((by, selector)))
        return True
    except:
        return False

def start_bpm():
    driver.get(url)

    login_field = pythonRPA.bySelector(
        [{"title": "WebDriver - Internet Explorer", "class_name": "Alternate Modal Top Most", "backend": "uia"},
         {"depth_start": 3, "depth_end": 3, "title": "User name", "control_type": "Edit"}])
    login_field.wait_appear(20)
    login_field.click()
    pythonRPA.keyboard.write(username)

    password_field = pythonRPA.bySelector(
        [{"title": "WebDriver - Internet Explorer", "class_name": "Alternate Modal Top Most", "backend": "uia"},
         {"depth_start": 3, "depth_end": 3, "title": "Password", "control_type": "Edit"}])
    password_field.click()
    pythonRPA.keyboard.write(password)

    pythonRPA.keyboard.press("enter")

    sleep(2)

    main_window = pythonRPA.bySelector([{"title": "Система управления кредитными заявками - Отчет - Internet Explorer",
                                         "class_name": "IEFrame", "backend": "uia"}])
    main_window.wait_appear(15)
    main_window.maximize()

    sleep(1)

def stop_bpm():
    driver.quit()
    #os.system("taskkill /f /im iexplore.exe")


def block_bpm(name):
    log_bpm_status = "Заблокирован"
    log_comments = ""

    #href = "Users.aspx"
    # WebDriverWait(driver, 20).until(
    #     EC.presence_of_all_elements_located((By.XPATH, "//a[@href = 'Site/Administration/Users.aspx']")))
    # users_menuOption = driver.find_element(By.XPATH, "//a[@href = 'Site/Administration/Users.aspx']")
    # users_menuOption.click()

    WebDriverWait(driver, 20).until(
        EC.presence_of_all_elements_located((By.XPATH, "//a[text() = 'Пользователи']")))
    users_menuOption = driver.find_element(By.XPATH, "//a[text() = 'Пользователи']")
    users_menuOption.click()

    sleep(0.5)

    WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input#TB_Search_FullName")))
    name_field = driver.find_element(By.CSS_SELECTOR, "input#TB_Search_FullName")
    name_field.send_keys(name)

    WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input#Button_Search")))
    search_button = driver.find_element(By.CSS_SELECTOR, "input#Button_Search")
    search_button.click()

    if __selenium_wait_element(By.CSS_SELECTOR, "table#GV_IpotekaUsers"):
        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "table#GV_IpotekaUsers")))
        results_table = driver.find_element(By.CSS_SELECTOR, "table#GV_IpotekaUsers")
        results_table_rows = results_table.find_elements(By.XPATH, ".//tr")
        print(len(results_table_rows))

        iterable = iter(results_table_rows)
        next(iterable)
        row_counter = 0
        for row in iterable:
            #print(row.get_attribute('innerHTML'))
            href_value = "javascript:__doPostBack('GV_IpotekaUsers','Select$" + str(row_counter) + "')"

            # WebDriverWait(driver, 10).until(
            #     EC.presence_of_all_elements_located((By.XPATH, ".//a")))
            # choose_button = row.find_element(By.XPATH, ".//a")    and text()='Выбрать'
            WebDriverWait(driver, 20).until(
                EC.presence_of_all_elements_located((By.XPATH, "//a[text() = 'Выбрать' and @href[contains(., 'Select$%s')]]"%row_counter)))
            choose_button = driver.find_element(By.XPATH, "//a[text() = 'Выбрать' and @href[contains(., 'Select$%s')]]"%row_counter)
            sleep(0.3)
            choose_button.click()
            sleep(0.2)
            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input#IB_Update")))
            edit_button = driver.find_element(By.CSS_SELECTOR, "input#IB_Update")
            edit_button.click()

            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input#TB_Nomad")))
            nomad_field = driver.find_element(By.CSS_SELECTOR, "input#TB_Nomad")
            nomad_field_value = nomad_field.get_attribute('value')
            print(len(nomad_field_value))
            if len(nomad_field_value) > 0:
                nomad_field.click()
                nomad_field.send_keys(Keys.CONTROL, "a")
                nomad_field.send_keys(Keys.DELETE)

            # x, y = pyautogui.locateCenterOnScreen(r"..\Utils\nomad_checkbox_chrome.png")
            # pythonRPA.mouse.click(x + 100, y)
            nomad_checkbox_clicked = pythonRPA.byImage(
                r".\Utils\nomad_checkbox.png")
            if nomad_checkbox_clicked.is_exists():
                print("True")
                WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input#NomadEnabled")))
                nomad_checkbox = driver.find_element(By.CSS_SELECTOR, "input#NomadEnabled")
                nomad_checkbox.click()
            else: print("False")

            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input#Button_DeleteExecutor")))
            delete_button_gruppy = driver.find_element(By.CSS_SELECTOR, "input#Button_DeleteExecutor")
            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input#Button_DeleteRole")))
            delete_button_roli = driver.find_element(By.CSS_SELECTOR, "input#Button_DeleteRole")

            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "select#LB_UserDepExecutors")))
            gruppy_ispolnitelei_area = driver.find_element(By.CSS_SELECTOR, "select#LB_UserDepExecutors")
            gruppy_ispolnitelei_items = gruppy_ispolnitelei_area.find_elements(By.XPATH, ".//option")
            print(len(gruppy_ispolnitelei_items))
            for item in range(len(gruppy_ispolnitelei_items)):
                driver.find_element(By.CSS_SELECTOR, "select#LB_UserDepExecutors").click()
                driver.find_element(By.CSS_SELECTOR, "select#LB_UserDepExecutors").send_keys(Keys.UP)
                sleep(0.5)
                driver.find_element(By.CSS_SELECTOR, "input#Button_DeleteExecutor").click()
                sleep(0.5)

            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "select#LB_UserRoles")))
            roli_area = driver.find_element(By.CSS_SELECTOR, "select#LB_UserRoles")
            roli_items = roli_area.find_elements(By.XPATH, ".//option")
            print(len(roli_items))
            for item in range(len(roli_items)):
                driver.find_element(By.CSS_SELECTOR, "select#LB_UserRoles").click()
                driver.find_element(By.CSS_SELECTOR, "select#LB_UserRoles").send_keys(Keys.UP)
                sleep(0.3)
                driver.find_element(By.CSS_SELECTOR, "input#Button_DeleteRole").click()
                sleep(0.3)

            sleep(0.5)
            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input#Button_AddRole")))
            driver.find_element(By.CSS_SELECTOR, "input#Button_AddRole").click()
            main_page = driver.current_window_handle
            #print(main_page)
            #print("----------------------------------------")
            sleep(3)
            for handle in driver.window_handles:
                if handle != main_page:
                    #print(handle)
                    add_role_window = handle
                    break
            driver.switch_to.window(add_role_window)

            # driver.switch_to.window(driver.window_handles[-1])
            # sleep(1)
            # print(driver.current_window_handle)

            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "select#DDL_PageSize")))
            select_button = driver.find_element(By.CSS_SELECTOR, "select#DDL_PageSize")
            select_button.click()
            sleep(0.5)
            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, ".//option[text()='Все']")))
            all_option = select_button.find_element(By.XPATH, ".//option[text()='Все']")
            all_option.click()
            sleep(0.5)
            roles_table = driver.find_element(By.CSS_SELECTOR, "table#GV_Roles")
            uchastnik_role_item = roles_table.find_element(By.XPATH, ".//tr[td[text()='Участник']]")
            select_role_button = uchastnik_role_item.find_element(By.XPATH, ".//a")
            select_role_button.click()
            sleep(0.5)
            driver.find_element(By.CSS_SELECTOR, "input#IB_Select").click()
            sleep(0.5)
            #driver.close()
            sleep(0.5)
            driver.switch_to.window(main_page)

            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input#Button_OK")))
            ok_button = driver.find_element(By.CSS_SELECTOR, "input#Button_OK")
            ok_button.click()

            row_counter += 1
    else:
        log_bpm_status = "Не заблокирован"
        log_comments = "Сотрудник не найден в системе"
    print("iteration ended")
    # for i in range(len(results_table_rows)-1):
    #     print("Hola"+str(i))
    #     choose_button
    # else:
    #     log_bpm_status = "Failed"
    #     log_comments = "Not found in system"
    #
    return log_bpm_status, log_comments


#
# start_bpm()
# for name in names:
#     print(block_bpm(name))
# stop_bpm()
