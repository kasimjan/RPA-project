import datetime
from datetime import datetime
from pyPythonRPA.Robot import pythonRPA
import pyautogui
import wmi
from pyautogui import*

import win32api
import win32con
import time

keys = {
    "left" : win32con.VK_LEFT,
    "right" : win32con.VK_RIGHT,
    "up" : win32con.VK_UP,
    "down" : win32con.VK_DOWN,
    "shift" : win32con.VK_SHIFT
}

def press_key(key, delay=0.1):
    if key in keys:
        win32api.keybd_event(keys[key], 0, win32con.KEYEVENTF_EXTENDEDKEY, 0)
        pythonRPA.keyboard.press("down", 10)
        time.sleep(delay)
        win32api.keybd_event(keys[key], 0, win32con.KEYEVENTF_KEYUP, 0)
    else:
        print("KEY NOT AVAILABLE")

press_key("shift", 5)
pythonRPA.keyboard.press("Shift")
def block(names,log_table):
    # start Active  Directory
    pythonRPA.keyboard.press('win + R')
    Run = pythonRPA.bySelector([{"title": "Run", "backend": "win32"}])
    Run.wait_appear(2)
    pythonRPA.keyboard.write('dsa.msc')
    pythonRPA.keyboard.press('Enter')

    # clicking the SEARCH button to open the search window
    AD_main = pythonRPA.bySelector(
        [{"title": "Active Directory Users and Computers", "class_name": "MMCMainFrame", "backend": "win32"}])
    AD_main.wait_appear(5)
    AD_main.set_focus()
    AD_main.maximize()
    print("Active Directory opened well!")

    #Action click
    pyautogui.click(62, 36)
    #change_domain click
    while 1:
        try:
            x, y = pyautogui.locateCenterOnScreen(r'.\Utils\change_domain.png')
            pyautogui.click(x, y)
            print('change domain clickd !')
            pythonRPA.sleep(2)
            break
        except Exception as e:
            print(e)
    pythonRPA.keyboard.press('Enter')
    pythonRPA.sleep(2)
    # # click hcsbk
    # pythonRPA.keyboard.press('down')
    # pythonRPA.sleep(2)
    # pythonRPA.keyboard.press('Enter')
    while 1:
        try:
            x, y = pyautogui.locateCenterOnScreen(r'.\Utils\hcsbk.png')
            pyautogui.doubleClick(x, y)
            pythonRPA.sleep(2)
            print('clicked hcsbk well !')
            break
        except Exception as e:
            print(e)

    while 1:
        try:
            x, y = pyautogui.locateCenterOnScreen(r'.\Utils\SEARCH.png')
            pyautogui.doubleClick(x, y)
            break
        except Exception as e:
            print(e)
    find_users = pythonRPA.bySelector(
        [{"title": "Find Users, Contacts, and Groups", "class_name": "#32770", "backend": "win32"}])
    find_users.wait_appear(2)
    find_users.set_focus()
    find_users.maximize()

    for name in names:
        try:
            # log_ad_status = "Success"
            # log_ad_comment = ""
            success = 0
            # searching the accounts by full names
            if name != None and name == 'test test test':
                search_by_name = pythonRPA.bySelector(
                    [{"title": "Find Users, Contacts, and Groups", "class_name": "#32770", "backend": "win32"},
                     {"ctrl_index": 152},{"ctrl_index": 7}])
                clear_button = pythonRPA.bySelector(
                    [{"title": "Find Users, Contacts, and Groups", "class_name": "#32770", "backend": "win32"},
                     {"ctrl_index": 14}])
                clear_button.click()
                pythonRPA.keyboard.press('Enter')
                search_by_name.wait_appear(1)
                search_by_name.click()
                pythonRPA.keyboard.write(name)
                pythonRPA.sleep(1)
                search_ok = pythonRPA.bySelector(
                    [{"title": "Find Users, Contacts, and Groups", "class_name": "#32770", "backend": "win32"},
                     {"ctrl_index": 12}])
                search_ok.click()
                pythonRPA.sleep(3)
                # clicking the result
                found = False
                for i in range(5):
                    try:
                        x,y = pyautogui.locateCenterOnScreen(r'.\Utils\found.PNG')
                        found = True
                        break
                    except Exception as e:
                        print(e)

                if found:
                    x = 178
                    y = 340
                    pyautogui.rightClick(x, y)

                    # make an account disable
                    for i in range(5):
                        try:
                            a,b = pyautogui.locateCenterOnScreen(r'.\Utils\enable_condition.png')
                            for i in range(4):
                                pythonRPA.keyboard.press('down')
                                print(a,b)
                            pythonRPA.keyboard.press('Enter')
                            pythonRPA.sleep(1)
                            pythonRPA.keyboard.press('Enter')
                            is_enable =  True
                            break
                        except: continue

                    #
                    success += 1
                    print(success)

                    # Удаляем члены групп
                    # открытие свойств
                    pyautogui.rightClick(x, y)
                    pythonRPA.keyboard.press('up arrow')
                    pythonRPA.sleep(1)
                    pythonRPA.keyboard.press('Enter')
                    pythonRPA.sleep(3)

                    name_variants = get_profile_name(name)
                    for i in range(len(name_variants)):
                        try:
                            user_profile = pythonRPA.bySelector(
                                [{"title":name_variants[i] + " Properties", "class_name": "#32770", "backend": "win32"}])
                            user_profile.set_focus()
                            name = name_variants[i]
                            break
                        except Exception as e:
                            print(e)

                    while 1:
                        try:
                            x, y = pyautogui.locateCenterOnScreen(r'.\Utils\member_of.png')
                            pyautogui.doubleClick(x, y)
                            break
                        except Exception as e:
                            print(e)
                    domain_sub_group = pythonRPA.byImage(r'.\Utils\domain_users.png')

                    n = 0

                    try:
                        sub_group = pythonRPA.bySelector(
                            [{"title": name + " Properties", "class_name": "#32770", "backend": "win32"},
                             {"ctrl_index": 17}, {"ctrl_index": 2}])
                        sub_group.wait_appear(1)
                        sub_group.click()
                        pythonRPA.keyboard.press('up')
                        press_key("shift", 5)
                        pythonRPA.keyboard.press("Shift")
                        remove_button = pythonRPA.bySelector(
                            [{"title": name + " Properties", "class_name": "#32770", "backend": "win32"},
                             {"ctrl_index": 0},
                             {"ctrl_index": 4}])
                        remove_button.click()
                        pythonRPA.keyboard.press('left')
                        pythonRPA.sleep(1)
                        pythonRPA.keyboard.press('Enter')
                        warning = pythonRPA.bySelector(
                            [{"title": "Active Directory Domain Services", "class_name": "#32770", "backend": "win32"}])
                        if (warning.is_exists()):
                            pythonRPA.keyboard.press('Enter')
                        ok = pythonRPA.bySelector([{"title":name+" Properties","class_name":"#32770","backend":"win32"},{"ctrl_index":12}])
                        ok.click()
                        pythonRPA.keyboard.press('Enter')
                    except Exception as e:
                        pythonRPA.keyboard.press('Enter')
                        break
                    success += 1
                    print(success)

                    # move this account to th blocked accounts
                    find_users.set_focus()
                    find_users.maximize()
                    clear_button = pythonRPA.bySelector(
                        [{"title": "Find Users, Contacts, and Groups", "class_name": "#32770", "backend": "win32"},
                         {"ctrl_index": 14}])
                    clear_button.click()
                    pythonRPA.keyboard.press('Enter')
                    search_by_name.wait_appear(1)
                    search_by_name.click()
                    pythonRPA.keyboard.write(name)
                    pythonRPA.sleep(1)
                    search_ok = pythonRPA.bySelector(
                        [{"title": "Find Users, Contacts, and Groups", "class_name": "#32770", "backend": "win32"},
                         {"ctrl_index": 12}])
                    search_ok.click()
                    pythonRPA.sleep(3)

                    # clicking the result
                    x = 178
                    y = 340
                    pyautogui.rightClick(x, y)
                    search_by_name.wait_appear(1)
                    search_by_name.click()
                    pyautogui.rightClick(x, y)
                    for i in range(6):
                        pythonRPA.keyboard.press('down arrow')
                    pythonRPA.keyboard.press('Enter')
                    pythonRPA.sleep(1)
                    move_to_block = pythonRPA.bySelector([{"title": "Move ", "class_name": "#32770", "backend": "win32"}])
                    move_to_block.wait_appear(1)
                    move_to_block.set_focus()
                    for i in range(10):
                        pythonRPA.keyboard.press('down arrow')
                    pythonRPA.keyboard.press('Enter')
                    print(name, 'Account added to BLOCK sucsessfully !')
                    pythonRPA.sleep(1)
                    success += 1
                    print(success)
                    if success == 3:
                        log_table.append([name, "Active Directory", "Заблокирован", "", datetime.now().strftime("%H:%M:%S")])
                    else:
                        log_table.append([name, "Active Directory", "Не заблокирован", "Сотрудник не найден в системе", datetime.now().strftime("%H:%M:%S")])
        except Exception as e:
            log_table.append([name, "Active Directory", "Не заблокирован", "Техническая ошибка", datetime.now().strftime("%H:%M:%S")])
    try:
        f = wmi.WMI()
        for p in f.Win32_Process():
         if p.name == 'mmc.exe':
             p.Terminate()
    except Exception as e:
        print(e)
def get_profile_name(name):
    name2 = name.split(' ')
    name_variants=[]
    result = ""
    for i in range(len(name2)):
        if i:
         result+=' '+name2[i]
        else: result+=name2[i]
        name_variants.append(result)
    print(name_variants)
    return name_variants


