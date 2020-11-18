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

options = Options()
options.binary_location = "C:\Program Files\Google\Chrome\Application\chrome.exe"
url = "http://bpmd/"
username = "mazhit.e"
password = "Asd_12345678"

driver = webdriver.Chrome(r"C:\Windows\System32\chromedriver.exe", chrome_options=options)

names = ['Амангельдиев Руслан Амангельдиевич', 'Абдрахманова Сандугаш Шекебаевна']


# names = ['Агимова Жансая Кудайбергеновна', 'Кантемирова Надежда Викторовна', 'Адилова Айгуль Сарсенбаевна', 'Абылкасимова Айгерим Сейтжановна']

def start_bpm():
    driver.get(url)

    login_form = pythonRPA.bySelector([{"title": "Sign in", "class_name": "Chrome_WidgetWin_1", "backend": "uia"},
                                       {"title": "Sign in", "control_type": "Custom"}])
    login_form.wait_appear(30)
    pythonRPA.keyboard.write(username)

    sleep(1)

    x, y = pyautogui.locateCenterOnScreen(r"..\Utils\password_field_title_bpm.png")
    pythonRPA.mouse.click(x + 100, y)
    # password_field_title = pythonRPA.byImage(r"..\Utils\password_field_title_bpm.png")
    # password_field_title.wait_appear(5)
    sleep(1)
    pythonRPA.keyboard.write(password)
    sleep(1)
    pythonRPA.keyboard.press("enter")
    pythonRPA.sleep(3)

    main_window = pythonRPA.bySelector(
        [{"title": "Система управления кредитными заявками - Отчет - Google Chrome", "class_name": "Chrome_WidgetWin_1",
          "backend": "uia"}])
    main_window.wait_appear(15)
    main_window.maximize()

def stop_bpm():
    driver.quit()


def after_choose_button_clicked():
    edit_button = pythonRPA.bySelector([{
        "title": "Система управления бизнес процессами - Пользователи - Internet Explorer",
        "class_name": "IEFrame", "backend": "uia"},
        {"depth_start": 16, "depth_end": 16, "title": "Изменить", "control_type": "Button"}])
    edit_button.click()

    gruppa_ispolnitelei_area = pythonRPA.bySelector([{
        "title": "Система управления бизнес процессами - Пользователи - Internet Explorer",
        "class_name": "IEFrame", "backend": "uia"}, {"ctrl_index": 2},
        {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 0},
        {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 2},
        {"ctrl_index": 0}, {"ctrl_index": 1}, {"ctrl_index": 2},
        {"ctrl_index": 3}, {"ctrl_index": 0}, {"ctrl_index": 7},
        {"ctrl_index": 1}, {"ctrl_index": 1},
        {"ctrl_index": 0, "control_type": "List"}])
    gruppa_ispolnitelei_area.wait_appear(5)
    if not gruppa_ispolnitelei_area.is_exists():
        gruppa_ispolnitelei_area = pythonRPA.bySelector([{
            "title": "Система управления бизнес процессами - Пользователи - Internet Explorer",
            "class_name": "IEFrame", "backend": "uia"}, {"ctrl_index": 3},
            {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 0},
            {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 2},
            {"ctrl_index": 0}, {"ctrl_index": 1}, {"ctrl_index": 2},
            {"ctrl_index": 3}, {"ctrl_index": 0}, {"ctrl_index": 7},
            {"ctrl_index": 1}, {"ctrl_index": 1},
            {"ctrl_index": 0, "control_type": "List"}])

    if len(gruppa_ispolnitelei_area.texts()) > 0:
        # print(len(gruppa_ispolnitelei_area.texts()))
        for i in range(len(gruppa_ispolnitelei_area.texts())):
            delete_button = pythonRPA.bySelector([{
                "title": "Система управления бизнес процессами - Пользователи - Internet Explorer",
                "class_name": "IEFrame", "backend": "uia"}, {"ctrl_index": 2},
                {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 0},
                {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 2},
                {"ctrl_index": 0}, {"ctrl_index": 1}, {"ctrl_index": 2},
                {"ctrl_index": 3}, {"ctrl_index": 0}, {"ctrl_index": 7},
                {"ctrl_index": 1}, {"ctrl_index": 2},
                {"title": "Удалить", "control_type": "Button"}])
            delete_button.wait_appear(5)
            if not delete_button.is_exists():
                delete_button = pythonRPA.bySelector([{
                    "title": "Система управления бизнес процессами - Пользователи - Internet Explorer",
                    "class_name": "IEFrame", "backend": "uia"}, {"ctrl_index": 3},
                    {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 0},
                    {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 2},
                    {"ctrl_index": 0}, {"ctrl_index": 1}, {"ctrl_index": 2},
                    {"ctrl_index": 3}, {"ctrl_index": 0}, {"ctrl_index": 7},
                    {"ctrl_index": 1}, {"ctrl_index": 2},
                    {"title": "Удалить", "control_type": "Button"}])

            gruppa_ispolnitelei_area.wait_appear(60)
            gruppa_ispolnitelei_area.click()
            pythonRPA.keyboard.press("up")
            delete_button.click()
            pythonRPA.mouse.click(400, 780)

    roli_area = pythonRPA.bySelector(
        [{"title": "Система управления бизнес процессами - Пользователи - Internet Explorer",
          "class_name": "IEFrame", "backend": "uia"}, {"ctrl_index": 2}, {"ctrl_index": 0},
         {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 0},
         {"ctrl_index": 2}, {"ctrl_index": 0}, {"ctrl_index": 1}, {"ctrl_index": 2},
         {"ctrl_index": 3}, {"ctrl_index": 0}, {"ctrl_index": 7}, {"ctrl_index": 2},
         {"ctrl_index": 1}, {"ctrl_index": 0, "control_type": "List"}])
    roli_area.wait_appear(10)
    if not roli_area.is_exists():
        roli_area = pythonRPA.bySelector(
            [{"title": "Система управления бизнес процессами - Пользователи - Internet Explorer",
              "class_name": "IEFrame", "backend": "uia"}, {"ctrl_index": 3}, {"ctrl_index": 0},
             {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 0},
             {"ctrl_index": 2}, {"ctrl_index": 0}, {"ctrl_index": 1}, {"ctrl_index": 2},
             {"ctrl_index": 3}, {"ctrl_index": 0}, {"ctrl_index": 7}, {"ctrl_index": 2},
             {"ctrl_index": 1}, {"ctrl_index": 0, "control_type": "List"}])
        roli_area.wait_appear(10)

    if len(roli_area.texts()) > 0:
        # print(len(roli_area.texts()))
        for i in range(len(roli_area.texts())):
            delete_button2 = pythonRPA.bySelector([{
                "title": "Система управления бизнес процессами - Пользователи - Internet Explorer",
                "class_name": "IEFrame", "backend": "uia"}, {"ctrl_index": 2},
                {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 0},
                {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 2},
                {"ctrl_index": 0}, {"ctrl_index": 1}, {"ctrl_index": 2},
                {"ctrl_index": 3}, {"ctrl_index": 0}, {"ctrl_index": 7},
                {"ctrl_index": 2}, {"ctrl_index": 2},
                {"title": "Удалить", "control_type": "Button"}])
            delete_button2.wait_appear(5)
            if not delete_button2.is_exists():
                delete_button2 = pythonRPA.bySelector([{
                    "title": "Система управления бизнес процессами - Пользователи - Internet Explorer",
                    "class_name": "IEFrame", "backend": "uia"}, {"ctrl_index": 3},
                    {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 0},
                    {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 2},
                    {"ctrl_index": 0}, {"ctrl_index": 1}, {"ctrl_index": 2},
                    {"ctrl_index": 3}, {"ctrl_index": 0}, {"ctrl_index": 7},
                    {"ctrl_index": 2}, {"ctrl_index": 2},
                    {"title": "Удалить", "control_type": "Button"}])
                delete_button2.wait_appear(5)

            roli_area.wait_appear(30)
            roli_area.click()
            pythonRPA.keyboard.press("up")
            delete_button2.click()
            pythonRPA.mouse.click(400, 780)

    # add_button = pythonRPA.bySelector(
    #     [{"title": "Система управления бизнес процессами - Пользователи - Internet Explorer",
    #       "class_name": "IEFrame", "backend": "uia"}, {"ctrl_index": 3}, {"ctrl_index": 0},
    #      {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 0},
    #      {"ctrl_index": 2}, {"ctrl_index": 0}, {"ctrl_index": 1}, {"ctrl_index": 2},
    #      {"ctrl_index": 3}, {"ctrl_index": 0}, {"ctrl_index": 7}, {"ctrl_index": 2},
    #      {"ctrl_index": 2}, {"title": "Добавить", "control_type": "Button"}])
    # add_button.wait_appear(15)
    # add_button.click()
    pythonRPA.mouse.click(220, 680)

    list_items_number = pythonRPA.bySelector(
        [{"title": "Список ролей - Internet Explorer", "class_name": "IEFrame", "backend": "uia"},
         {"depth_start": 15, "depth_end": 15, "title": "Open", "control_type": "Button"}])
    list_items_number.wait_appear(30)
    list_items_number.click()
    sleep(1)
    pythonRPA.keyboard.press("down", 5)
    pythonRPA.keyboard.press("enter")
    pythonRPA.mouse.scroll(6)

    roles_list = pythonRPA.bySelector(
        [{"title": "Список ролей - Internet Explorer", "class_name": "IEFrame", "backend": "uia"}])
    roles_list.click()
    pythonRPA.keyboard.press("down", 40)

    # member_menu_item = pythonRPA.bySelector()
    member_menu_item = pythonRPA.byImage(
        r".\Utils\word_member.png")
    member_menu_item.wait_appear(15)
    # print(member_menu_item.attrs())
    attrs = member_menu_item.attrs()
    left = attrs.left
    top = attrs.top
    # print(left)
    # print(top)
    pythonRPA.mouse.click(left - 30, top + 5)
    # member_menu_item.near( near_image_path=r"C:\Users\mazhit.e\Desktop\ttt\PythonRPA-master\Robot\Blocking_Robot\Utils\href_select.png",position="right", distance_in_px=20)

    # select_member_menu_item.click()

    final_select_button = pythonRPA.bySelector(
        [{"title": "Список ролей - Internet Explorer", "class_name": "IEFrame", "backend": "uia"},
         {"depth_start": 14, "depth_end": 14, "title": "Выбрать", "control_type": "Button"}])
    final_select_button.click()
    sleep(2)

    pythonRPA.mouse.click(550, 370)
    pythonRPA.keyboard.press("ctrl+a")
    pythonRPA.keyboard.press("delete")

    sleep(2)

    nomad_checkbox_clicked = pythonRPA.byImage(
        r".\Utils\nomad_checkbox.png")
    if nomad_checkbox_clicked.is_exists():
        nomad_checkbox = pythonRPA.bySelector([{
            "title": "Система управления бизнес процессами - Пользователи - Internet Explorer",
            "class_name": "IEFrame", "backend": "uia"},
            {"depth_start": 16, "depth_end": 16, "title": "", "control_type": "CheckBox"}])
        nomad_checkbox.click()

    sleep(2)

    ok_button = pythonRPA.bySelector(
        [{"title": "Система управления бизнес процессами - Пользователи - Internet Explorer",
          "class_name": "IEFrame", "backend": "uia"},
         {"depth_start": 16, "depth_end": 16, "title": "OK", "control_type": "Button"}])
    ok_button.click()


def block_bpm(name):
    log_bpm_status = "Success"
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

    WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input#TB_Search_FullName")))
    name_field = driver.find_element(By.CSS_SELECTOR, "input#TB_Search_FullName")
    name_field.send_keys(name)

    WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input#Button_Search")))
    search_button = driver.find_element(By.CSS_SELECTOR, "input#Button_Search")
    search_button.click()

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
            EC.presence_of_all_elements_located((By.XPATH, "//a[@href[contains(., 'Select$%s')]]"%row_counter)))
        choose_button = driver.find_element(By.XPATH, "//a[@href[contains(., 'Select$%s')]]"%row_counter)
        sleep(1)
        choose_button.click()
        sleep(1)
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
            r"..\Utils\nomad_checkbox_chrome.png")
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
            sleep(1)

        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "select#LB_UserRoles")))
        roli_area = driver.find_element(By.CSS_SELECTOR, "select#LB_UserRoles")
        roli_items = roli_area.find_elements(By.XPATH, ".//option")
        print(len(roli_items))
        for item in range(len(roli_items)):
            driver.find_element(By.CSS_SELECTOR, "select#LB_UserRoles").click()
            driver.find_element(By.CSS_SELECTOR, "select#LB_UserRoles").send_keys(Keys.UP)
            sleep(0.5)
            driver.find_element(By.CSS_SELECTOR, "input#Button_DeleteRole").click()
            sleep(1)

        sleep(2)
        driver.find_element(By.CSS_SELECTOR, "input#Button_AddRole").click()


        sleep(3)

        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input#Button_OK")))
        ok_button = driver.find_element(By.CSS_SELECTOR, "input#Button_OK")
        ok_button.click()




        row_counter += 1


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
start_bpm()
for name in names:
    print(block_bpm(name))
stop_bpm()
