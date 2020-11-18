from pyPythonRPA.Robot import pythonRPA
from selenium import webdriver
import os

from time import sleep
from selenium.webdriver.chrome.options import Options

# options = Options()
# options.binary_location = "C:\Program Files\Google\Chrome\Application\chrome.exe"

# names = get_names()
# for name in names:

username = "mazhit.e"
password = "Asd_12345678"
url = "http://bpmd/"
# names = ['Муканов Талгат Лаикович', 'Серик Данияр']
#names = ['Агимова Жансая Кудайбергеновна', 'Кантемирова Надежда Викторовна', 'Адилова Айгуль Сарсенбаевна', 'Абылкасимова Айгерим Сейтжановна']

driver = webdriver.Ie(
    r"C:\Windows\System32\IEDriverServer.exe")


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


def stop_bpm():
    driver.quit()
    os.system("taskkill /f /im iexplore.exe")


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

    users_menuOption = pythonRPA.bySelector(
        [{"title": "Система управления кредитными заявками - Отчет - Internet Explorer",
          "class_name": "IEFrame", "backend": "uia"},
         {"depth_start": 12, "depth_end": 12, "title": "Пользователи", "control_type": "Hyperlink"}])
    users_menuOption.wait_appear(20)

    if users_menuOption.is_exists():
        users_menuOption.click()
    else:
        users_menuOption = pythonRPA.bySelector(
            [{"title": "Система управления бизнес процессами - Пользователи - Internet Explorer",
              "class_name": "IEFrame",
              "backend": "uia"},
             {"depth_start": 12, "depth_end": 12, "title": "Пользователи", "control_type": "Hyperlink"}])
        users_menuOption.wait_appear(20)
        users_menuOption.click()

    sleep(3)

    name_field = pythonRPA.bySelector(
        [{"title": "Система управления бизнес процессами - Пользователи - Internet Explorer",
          "class_name": "IEFrame", "backend": "uia"}, {"ctrl_index": 2}, {"ctrl_index": 0},
         {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 0},
         {"ctrl_index": 2}, {"ctrl_index": 0}, {"ctrl_index": 1}, {"ctrl_index": 2},
         {"ctrl_index": 3}, {"ctrl_index": 0}, {"ctrl_index": 7}, {"ctrl_index": 0},
         {"ctrl_index": 3}, {"ctrl_index": 0}])
    name_field.wait_appear(10)

    if name_field.is_exists():
        name_field.click()
    else:
        name_field = pythonRPA.bySelector(
            [{"title": "Система управления бизнес процессами - Пользователи - Internet Explorer",
              "class_name": "IEFrame", "backend": "uia"}, {"ctrl_index": 3}, {"ctrl_index": 0},
             {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 0},
             {"ctrl_index": 2}, {"ctrl_index": 0}, {"ctrl_index": 1}, {"ctrl_index": 2},
             {"ctrl_index": 3}, {"ctrl_index": 0}, {"ctrl_index": 7}, {"ctrl_index": 0},
             {"ctrl_index": 3}, {"ctrl_index": 0}])
        name_field.wait_appear(15)
        name_field.click()

    pythonRPA.keyboard.write(name)

    search_button = pythonRPA.bySelector([{
        "title": "Система управления бизнес процессами - Пользователи - Internet Explorer",
        "class_name": "IEFrame", "backend": "uia"},
        {"depth_start": 16, "depth_end": 16, "title": "Поиск", "control_type": "Button"}])
    search_button.click()

    # choose_button = pythonRPA.bySelector([{
    #     "title": "Система управления бизнес процессами - Пользователи - Internet Explorer",
    #     "class_name": "IEFrame", "backend": "uia"},
    #     {"depth_start": 14, "depth_end": 14, "title": "Выбрать", "control_type": "Hyperlink"}])

    sleep(5)
    pythonRPA.mouse.click(100, 800)


    choose_button1_1 = pythonRPA.bySelector([{
        "title": "Система управления бизнес процессами - Пользователи - Internet Explorer",
        "class_name": "IEFrame", "backend": "uia"}, {"ctrl_index": 2},
        {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 0},
        {"ctrl_index": 0}, {"ctrl_index": 2}, {"ctrl_index": 0}, {"ctrl_index": 1},
        {"ctrl_index": 3}, {"ctrl_index": 3}, {"ctrl_index": 0}, {"ctrl_index": 12},
        {"title": "Выбрать", "control_type": "Hyperlink"}])

    choose_button1_2 = pythonRPA.bySelector(
        [{"title": "Система управления бизнес процессами - Пользователи - Internet Explorer", "class_name": "IEFrame",
          "backend": "uia"}, {"ctrl_index": 3}, {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 0},
         {"ctrl_index": 0},
         {"ctrl_index": 0}, {"ctrl_index": 2}, {"ctrl_index": 0}, {"ctrl_index": 1}, {"ctrl_index": 3},
         {"ctrl_index": 3},
         {"ctrl_index": 0}, {"ctrl_index": 12}, {"title": "Выбрать", "control_type": "Hyperlink"}])

    # choose_button2_1 = pythonRPA.bySelector(
    #     [{"title": "Система управления бизнес процессами - Пользователи - Internet Explorer", "class_name": "IEFrame",
    #       "backend": "uia"}, {"ctrl_index": 2}, {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 0},
    #      {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 2}, {"ctrl_index": 0}, {"ctrl_index": 1},
    #      {"ctrl_index": 3}, {"ctrl_index": 3}, {"ctrl_index": 0}, {"ctrl_index": 24},
    #      {"title": "Выбрать", "control_type": "Hyperlink"}])
    #
    # choose_button2_2 = pythonRPA.bySelector(
    #     [{"title": "Система управления бизнес процессами - Пользователи - Internet Explorer", "class_name": "IEFrame",
    #       "backend": "uia"}, {"ctrl_index": 3}, {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 0},
    #      {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 2}, {"ctrl_index": 0}, {"ctrl_index": 1},
    #      {"ctrl_index": 3}, {"ctrl_index": 3}, {"ctrl_index": 0}, {"ctrl_index": 24},
    #      {"title": "Выбрать", "control_type": "Hyperlink"}])
    #
    # choose_button3_1 = pythonRPA.bySelector(
    #     [{"title": "Система управления бизнес процессами - Пользователи - Internet Explorer", "class_name": "IEFrame",
    #       "backend": "uia"}, {"ctrl_index": 2}, {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 0},
    #      {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 2}, {"ctrl_index": 0}, {"ctrl_index": 1},
    #      {"ctrl_index": 3}, {"ctrl_index": 3}, {"ctrl_index": 0}, {"ctrl_index": 36},
    #      {"title": "Выбрать", "control_type": "Hyperlink"}])
    #
    # choose_button3_2 = pythonRPA.bySelector(
    #     [{"title": "Система управления бизнес процессами - Пользователи - Internet Explorer", "class_name": "IEFrame",
    #       "backend": "uia"}, {"ctrl_index": 3}, {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 0},
    #      {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 2}, {"ctrl_index": 0}, {"ctrl_index": 1},
    #      {"ctrl_index": 3}, {"ctrl_index": 3}, {"ctrl_index": 0}, {"ctrl_index": 36},
    #      {"title": "Выбрать", "control_type": "Hyperlink"}])

    result_table = pythonRPA.bySelector(
        [{"title": "Система управления бизнес процессами - Пользователи - Internet Explorer",
          "class_name": "IEFrame", "backend": "uia"}, {"depth_start": 13, "depth_end": 13, "title": "", "control_type": "HeaderItem"}])
    result_table.wait_appear(10)
    if result_table.is_exists():
        if choose_button1_1.is_exists():
            choose_button1_1.click()
            after_choose_button_clicked()
        else:
            choose_button1_2.click()
            after_choose_button_clicked()

        # if choose_button2_1.is_exists():
        #     choose_button2_1.click()
        #     after_choose_button_clicked()
        # elif choose_button2_2.is_exists():
        #     choose_button2_2.click()
        #     after_choose_button_clicked()
        #
        # if choose_button3_1.is_exists():
        #     choose_button3_1.click()
        #     after_choose_button_clicked()
        # elif choose_button3_2.is_exists():
        #     choose_button3_2.click()
        #     after_choose_button_clicked()
    else:
        log_bpm_status = "Failed"
        log_comments = "Not found in system"

    return log_bpm_status, log_comments

#
# start_bpm()
# for name in names:
#     print(block_bpm(name))
# stop_bpm()
