from pyPythonRPA.Robot import pythonRPA
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from time import sleep
from selenium.webdriver.common.action_chains import ActionChains

from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException


options = Options()
options.binary_location = "C:\Program Files\Google\Chrome\Application\chrome.exe"
url = "http://doc.hcsbk.kz/"
username = "robot_ad"
password = "Asd12345678"
#name = "Ахмет Мұрат Жаңбырбайұлы"
# names = ['Агимова Жансая Кудайбергеновна', 'Кантемирова Надежда Викторовна', 'Адилова Айгуль Сарсенбаевна', 'Абылкасимова Айгерим Сейтжановна']

driver = webdriver.Chrome(
    r"C:\Windows\System32\chromedriver.exe",
    chrome_options=options)

def reload_driver():
    global driver
    driver.close()
    driver = webdriver.Chrome(
        r"C:\Windows\System32\chromedriver.exe",
        chrome_options=options)

def __selenium_wait_element(by, selector, sec = 120, appear=True):
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


def start_documentolog():
    driver.get(url)

    main_window = pythonRPA.bySelector([{"title":"Вход — Documentolog 7 - Google Chrome","class_name":"Chrome_WidgetWin_1","backend":"uia"}])
    main_window.wait_appear(10)
    if not main_window.is_exists():
        main_window = pythonRPA.bySelector([{"title":"Документы - Google Chrome","class_name":"Chrome_WidgetWin_1","backend":"uia"}])
    main_window.wait_appear(10)
    main_window.maximize()

    WebDriverWait(driver, 15).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input#login")))
    login_field = driver.find_element(By.CSS_SELECTOR, "input#login")
    login_field.send_keys(username)
    WebDriverWait(driver, 5).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input#password")))
    password_field = driver.find_element(By.CSS_SELECTOR, "input#password")
    password_field.send_keys(password)
    WebDriverWait(driver, 5).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input#submit")))
    login_button = driver.find_element(By.CSS_SELECTOR, "input#submit")
    login_button.click()

    WebDriverWait(driver, 15).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "ul#header_menu")))
    #parent_menu = driver.find_element(By.CSS_SELECTOR, "div#header")
    parent_menu2 = driver.find_element(By.CSS_SELECTOR, "ul#header_menu")
    WebDriverWait(driver, 15).until(EC.presence_of_all_elements_located((By.XPATH, ".//li[a[text()[contains(., 'Справочники')]]]")))
    menu_item = parent_menu2.find_element(By.XPATH, ".//li[a[text()[contains(., 'Справочники')]]]")
    menu_item.click()

    WebDriverWait(driver, 15).until(EC.presence_of_all_elements_located((By.XPATH, ".//div/ul/li")))
    menu_items_next = menu_item.find_elements(By.XPATH, ".//div/ul/li")
    for item in menu_items_next:
        if "Структура" in item.text:
            item.click()
            break


    # menu_item_next = menu_item.find_element(By.XPATH, ".//div/ul/li[a[text() = 'Структура']]")
    # menu_item_next.click()

    # menu_items = parent_menu.find_elements(By.XPATH, "//ul[@id='header_menu']/li")
    # for menu_item in menu_items:
    #     if "Справочники" in menu_item.get_attribute('inner_HTML'):
    #         print("Yes")
    #     else:
    #         print("No")


def stop_documentolog():
    driver.quit()


def block_documentolog(name):
    log_documentolog_status = "Заблокирован"
    log_comments = ""
    WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input#structure_search_input")))
    search_field = driver.find_element(By.CSS_SELECTOR, "input#structure_search_input")
    search_field.send_keys(name)
    sleep(5)
    search_field.send_keys(Keys.ARROW_DOWN)
    search_field.send_keys(Keys.ENTER)

    WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div#structure_view")))
    structure_view = driver.find_element(By.CSS_SELECTOR, "div#structure_view")
    sleep(3)

    results_list = driver.find_element_by_class_name(
        'ui-autocomplete.ui-menu.ui-widget.ui-widget-content.ui-corner-all')
    results_count = results_list.find_elements(By.XPATH, './/li')
    # print(len(results_count))

    if len(results_count) == 1:
        log_documentolog_status = "Не заблокирован"
        log_comments = "Сотрудник не найден в системе"
        return log_documentolog_status, log_comments

    if __selenium_wait_element(By.XPATH, "//p[@class='selected']/span/span"):
        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, "//p[@class='selected']/span/span")))
        edit_button = structure_view.find_element(By.XPATH, "//p[@class='selected']/span/span")
        edit_button.click()

        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, "//p[@class='selected']/span/span/span/a[text()='Редактировать']")))
        edit_menu_item = structure_view.find_element(By.XPATH, "//p[@class='selected']/span/span/span/a[text()='Редактировать']")
        edit_menu_item.click()

        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "form#structureEditForm")))
        structure_edit_window = driver.find_element(By.CSS_SELECTOR, "form#structureEditForm")

        # WebDriverWait(driver, 20).until(
        #     EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div#tabcontent1")))
        # edit_form = driver.find_element(By.CSS_SELECTOR, "div#tabcontent1")
        # edit_form_fields = edit_form.find_elements(By.XPATH, './/div/div[div[text()]]/div')
        # #edit_form_fields = structure_edit_window.find_elements(By.XPATH, "// div[ @ id = 'tabcontent1'] / div / div / div")
        # for field in edit_form_fields:
        #     try:
        #         if field.text == "Название":
        #             title_field = field.find_element(By.XPATH, ".//input")
        #             title_field.click()
        #             title_field.send_keys(Keys.CONTROL, "a")
        #             title_field.send_keys(Keys.DELETE)
        #             title_field.send_keys("Заблокирован")
        #             break
        #     except StaleElementReferenceException:
        #         pass

        # input = driver.find_element(By.CSS_SELECTOR, "input#edit_field_display_name")
        # input.click()
        # input.send_keys(Keys.CONTROL, "a")
        # input.send_keys(Keys.DELETE)
        # input.send_keys("Заблокирован")
        #print(driver.find_element(By.CSS_SELECTOR, "div#tabcontent1").get_attribute('innerHTML'))
        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div#tabcontent1")))
        driver.find_element(By.CSS_SELECTOR, "div#tabcontent1").find_element(By.XPATH, ".//div/div[div[text()='Название']]").find_element(By.XPATH, ".//input").click()
        driver.find_element(By.CSS_SELECTOR, "div#tabcontent1").find_element(By.XPATH, ".//div/div[div[text()='Название']]").find_element(By.XPATH, ".//input").send_keys(Keys.CONTROL, "a")
        driver.find_element(By.CSS_SELECTOR, "div#tabcontent1").find_element(By.XPATH, ".//div/div[div[text()='Название']]").find_element(By.XPATH, ".//input").send_keys(Keys.DELETE)
        driver.find_element(By.CSS_SELECTOR, "div#tabcontent1").find_element(By.XPATH, ".//div/div[div[text()='Название']]").find_element(By.XPATH, ".//input").send_keys("Заблокирован")


        # for field in edit_form_fields:
        #     print(field.text)

        # wnd = structure_edit_window.find_element(By.XPATH, "//div[@id = 'tabcontent1' and div[div[text()='Информация должности']]/div[@class='panel-table']]")
        # wnd2 = wnd.find_element(By.XPATH, "//div[div[text()='Родительский элемент']]")
        # wnd3 = wnd2.find_elements(By.TAG_NAME, 'option')
        # print(len(wnd3))
        # #print(wnd2.get_attribute('innerHTML'))

        sleep(3)
        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div#tabcontent1")))
        wnd_up = driver.find_element(By.CSS_SELECTOR, "div#tabcontent1")
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, ".//div/div[div[text()]]/div")))
            wnd = wnd_up.find_elements(By.XPATH, ".//div/div[div[text()]]/div")
        except StaleElementReferenceException:
            pass
        try:
            for each in wnd:
                if "Родительский элемент" in each.text:
                    #print(each.get_attribute('innerHTML'))
                    div_input_field = each.find_element(By.XPATH, ".//div[a]")
                    #div_input_field2 = div_input_field.find_element(By.XPATH, ".//div[input]")
                    # div_input_field2 = div_input_field.find_element(By.XPATH, ".//div")
                    # print(div_input_field2.get_attribute('innerHTML'))
                    # sleep(1)
                    # div_input_field.click()
                    # sleep(1)
                    # input_field = div_input_field2.find_element(By.XPATH, ".//input")
                    WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input#structureAutocomplete_parent")))
                    input_field = driver.find_element(By.CSS_SELECTOR, "input#structureAutocomplete_parent")
                    #parent_id
                    input_field.click()
                    input_field.send_keys("Заблокированные")
                    sleep(1)
                    input_field.send_keys(Keys.ARROW_DOWN)
                    sleep(0.5)
                    input_field.send_keys(Keys.ENTER)
                    break
        except StaleElementReferenceException:
            pass


        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, ".//div[div[text()[contains(., 'Группы')]]]")))
        groups_edit_field = structure_edit_window.find_element(By.XPATH, ".//div[div[text()[contains(., 'Группы')]]]")
        #print(groups_edit_field.get_attribute('innerHTML'))
        groups_field_values = groups_edit_field.find_elements(By.XPATH, ".//ul/li")
        for item in groups_field_values:
            delete_groups_value = item.find_element(By.XPATH, ".//span[@class='remove_item_btn changemonitor']")
            delete_groups_value.click()

        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'li#tabtitle3')))
        tab_User = driver.find_element(By.CSS_SELECTOR, 'li#tabtitle3')
        tab_User.click()

        sleep(1)

        blocked_toggle_button = structure_edit_window.find_element(By.XPATH, ".// div[ @ id = 'tabcontent3'] / div/div/div/div/span/label[input[@value='Блокированный']]")
        blocked_toggle_button.click()


        sleep(2)
        # buttons = driver.find_elements_by_class_name("btn-approve.ui-button.ui-widget.ui-state-default.ui-corner-all.ui-button-text-only")
        # print (len(buttons))
        button = driver.find_element_by_class_name("btn-approve.ui-button.ui-widget.ui-state-default.ui-corner-all.ui-button-text-only")
        button.click()
        # print(test_form.get_attribute("innerHTML"))
        # main_edit_window = driver.find_elements(By.XPATH, "//div")
        # try:
        #     for i in main_edit_window:
        #
        #         if "РЕДАКТИРОВАНИЕ ДОЛЖНОСТИ" in i.text:
        #             if i.get_attribute("class") == "ui-dialog ui-widget ui-widget-content ui-corner-all":
        #                 button_set = i.find_element(By.XPATH, ".//div[@class='ui-dialog-buttonset']")
        #                 button = button_set.find_element(By.XPATH, ".//button[text()='OK']")
        #                 button.click()
        #                 # for button in buttons:
        #                 #     if button.text == "ОК":
        #                 #         button.click()
        #                 #         break
        # except StaleElementReferenceException:
        #     pass
    else:
        log_documentolog_status = "Не заблокирован"
        log_comments = "Сотрудник не найден в системе"

    return log_documentolog_status, log_comments


# for name in names:
#     start_documentolog()
#     print(block_documentolog(name))
#     sleep(2)
#     reload_driver()
# stop_documentolog()

#start_documentolog()
#block_documentolog(name)
#stop_documentolog()

