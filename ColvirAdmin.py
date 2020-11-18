import datetime
import os
from pyPythonRPA.Robot import pythonRPA
import pyautogui
from datetime import datetime
def change_row(log_table,row1,name,system,status,comment,time):
    for row in log_table:
        if row==row1:
            row[0] = name
            row[1] = system
            row[2] = status
            row[3] = comment
            row[4] = time
def ColvirAdmin_block(names,log_table):
    n = 2
    for i in range(n):
            # open ColvirAdmin
            login_cbs_adm = "arnurt"
            password_cbs_adm = "arnur_010203"
            pythonRPA.application("C:\CBS_T_новый\CBSADM.exe").start()

            #Введение данных для входа
            log_in = pythonRPA.bySelector([{"title": "Вход в систему", "class_name": "TfrmLoginDlg", "backend": "uia"}])
            log_in.wait_appear(30)
            log_in.click()
            log_in.set_focus()
            pythonRPA.keyboard.write(login_cbs_adm)
            pythonRPA.keyboard.press("Enter")
            pythonRPA.keyboard.write(password_cbs_adm)
            pythonRPA.keyboard.press('Enter')

            cbs_adm_main_window = pythonRPA.bySelector([{"class_name": "TfrmCssApplAdm", "backend": "uia"}])
            cbs_adm_main_window.wait_appear(10)
            cbs_adm_main_window.set_focus()

            #Сохранение селекторов Колвир Админ
            filter = pythonRPA.bySelector([{"title": "Фильтр", "class_name": "TfrmFilterParams", "backend": "uia"}])
            filter_name_line = pythonRPA.bySelector([{"title": "Фильтр", "class_name": "TfrmFilterParams", "backend": "uia"},
                                                     {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 0}])

            filter_cbs_ok = pythonRPA.bySelector([{"title": "Фильтр", "class_name": "TfrmFilterParams", "backend": "uia"},
                                                  {"ctrl_index": 1}, {"ctrl_index": 0}, {"ctrl_index": 2}])
            for name in names:

                # openning the filter by cliking the Пользователи
                try:
                    blocked_codes=[]
                    unblocked_codes = []
                    for iterr in range(50):
                        try:
                            x, y = pyautogui.locateCenterOnScreen(r'.\Utils\users.png')
                            pyautogui.doubleClick(x, y)
                            print('Пользователи pressed')
                            break
                        except Exception as e:
                            print(e)
                            pythonRPA.sleep(1)
                    pythonRPA.sleep(2)
                    filter.set_focus()

                    # Clearing the FILTER
                    for iterr in range(50):
                        try:
                            x, y = pyautogui.locateCenterOnScreen(r'.\Utils\clear_filter_button.png')
                            pyautogui.doubleClick(x, y)
                            print('clear_filter_button pressed')
                            break
                        except Exception as e:
                            print(e)
                            pythonRPA.sleep(1)
                    filter_name_line.click()
                    # fill name_line
                    pythonRPA.keyboard.write(name)

                    # clicking OK to search
                    filter_cbs_ok.click()
                    pythonRPA.sleep(3)
                    code = "npos"
                    #Проверка на результат
                    warming = pythonRPA.bySelector([{"title":"Подтверждение","class_name":"TMessageForm","backend":"uia"}])
                    warming.wait_appear(1)
                    unfound =  warming.is_exists()
                    if unfound:
                        if i:
                            for row in log_table:
                             if row[0]==name and row[1] == 'CBSADM':
                                 change_row(log_table,row,name, "CBSADM", "Не заблокирован", "Сотрудник не найден в системе",datetime.now().strftime("%H:%M:%S"))
                        else:
                            log_table.append([name, "CBSADM", "Не заблокирован", "Сотрудник не найден в системе",datetime.now().strftime("%H:%M:%S")])
                        print("Not found in system")
                        pythonRPA.keyboard.press('enter')
                        continue
                    else:
                        used  = False
                        while 1:
                            # Getting the CODE of the user
                            # i = 1
                            if(used):
                                pythonRPA.keyboard.press('down')
                            pythonRPA.keyboard.press('Enter')
                            used = True
                            user_detail_window = pythonRPA.bySelector([{"class_name": "TfrmAdmUsrDetail", "backend": "uia"}])
                            user_detail_window.wait_appear(2)
                            CODE_line = pythonRPA.bySelector([{"class_name": "TfrmAdmUsrDetail", "backend": "uia"}, {"ctrl_index": 0},
                                                              {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 2},
                                                              {"ctrl_index": 0}])

                            # check do we blocked this account in past itteration??
                            if (CODE_line.get_value() == code):
                                break

                            code = CODE_line.get_value()
                            # Флаг проеверки должна ли программа нажать кнопку архив, изначально думаем что должна
                            archive_сlick = True

                            # Флаг проеверки должна ли программа нажать кнопку блокировки, изначально думаем что должна
                            block_click =True

                            try:
                                x, y = pyautogui.locateCenterOnScreen(r'.\Utils\archive.png')
                                archive_сlick = False
                                # Флаг проеверки становится ложной если он и так нажат
                            except Exception as e:
                                print(e)
                                n+=1
                                print('archive ', archive_сlick)

                            try:
                                x, y = pyautogui.locateCenterOnScreen(r'.\Utils\blocked.png')
                                block_click = False
                                print('block_click ', block_click)
                                # Флаг проеверки становится ложной если он и так нажат
                            except Exception as e:
                                print(e)
                                print('block_click ', block_click)
                            # Кнопка "Пользователь"
                            user_button = pythonRPA.bySelector(
                                [{"class_name": "TfrmAdmUsrDetail", "backend": "uia"}, {"ctrl_index": 4}, {"ctrl_index": 0}])
                            # В сулачае если галочки нету или же кнопка блокировки не нажата мы должны ввести изминения
                            if (archive_сlick):
                                # Нажимаем кнопку корректировки
                                user_button.wait_appear(2)
                                user_button.click()
                                for i in range(3):
                                    pythonRPA.keyboard.press('down')
                                pythonRPA.keyboard.press('Enter')

                                # making enable the checkbox of the archive
                                try:
                                    x, y = pyautogui.locateCenterOnScreen(r'.\Utils\arxiv.png')
                                    pyautogui.click(x + (x / 2), y)
                                    print('Архив нажат')
                                except Exception as e:
                                    print(e)
                                # Saving the changes
                                user_button.wait_appear(2)
                                user_button.click()
                                for i in range(2):
                                    pythonRPA.keyboard.press('down')
                                pythonRPA.keyboard.press('Enter')
                             # Blocking the user
                            if block_click:
                                user_button.wait_appear(2)
                                user_button.click()
                                for i in range(9):
                                    pythonRPA.keyboard.press('down')
                                for i in range(3):
                                    pythonRPA.sleep(0.5)
                                    pythonRPA.keyboard.press('Enter')
                            user_detail_window.close()
                            print(code, "sucsessfully blocked in CBS/ADM.exe")
                    if not i:
                        log_table.append([name, "CBSADM","Заблокирован", "", datetime.now().strftime("%H:%M:%S")])
                    else:
                        for row in log_table:
                            if row[0] == name and row[1] == 'CBSADM':
                                change_row(log_table,row, name,"CBSADM","Заблокирован", "", datetime.now().strftime("%H:%M:%S"))
                except:
                    if i:
                        for row in log_table:
                            if row[0] == name and row[1] == 'CBSADM':
                                change_row(log_table,row, name,"CBSADM", "Не заблокирован", "Техническая ошибка", datetime.now().strftime("%H:%M:%S"))
                    else:
                        log_table.append([name, "CBSADM", "Не заблокирован", "Техническая ошибка", datetime.now().strftime("%H:%M:%S")])