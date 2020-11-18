import time
import os
from pyPythonRPA.Robot import pythonRPA
from datetime import  date
import datetime
import pyautogui
from datetime import datetime
def download_report():
    # while (1):
    #     print('start')
    #     if datetime.datetime.now().strftime("%H:%M") == "19:00" or 1:
    #         try:
    #             # запуск Colvirtest (40TP)
    #             pythonRPA.application("C:\CBS_T/COLVIR.EXE").start()
    #         except Exception as e:
    #             print('colvir error ' + e)
    #     break

    try:
    # запуск Colvirtest (40TP)
        pythonRPA.application("C:\CBS_T/COLVIR.EXE").start()
    except Exception as e:
        print('colvir error ' + e)
    # Выполнение входа
    log_in = pythonRPA.bySelector([{"title": "Вход в систему", "class_name": "TfrmLoginDlg", "backend": "uia"}])
    log_in.wait_appear(30)
    log_in.click()
    log_in.set_focus()
    pythonRPA.keyboard.write("colvir")
    pythonRPA.keyboard.press("Enter")
    pythonRPA.keyboard.write("test_2722")
    pythonRPA.keyboard.press('Enter')
    # ert_1908
    # konakayev_t

    # Перенаправление фокус экрана на выбор режима
    mode_choosing = pythonRPA.bySelector([{"title": "Выбор режима", "class_name": "TfrmCssMenu", "backend": "win32"}])
    mode_choosing.wait_appear(10)
    mode_choosing.set_focus()
    time.sleep(2)
    mode_s = pythonRPA.bySelector(
        [{"title": "Выбор режима", "class_name": "TfrmCssMenu", "backend": "win32"}, {"ctrl_index": 3},
         {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 0}])
    mode_s.click()

    # Очистка фильтра задач
    pythonRPA.keyboard.press('SHIFT+CTRL+A')
    pythonRPA.keyboard.press('Delete')
    pythonRPA.keyboard.write('PRS')
    pythonRPA.keyboard.press('Enter')

    # очистка фильтра
    pythonRPA.sleep(4)
    for iterr in range(50):
        try:
            x, y = pyautogui.locateCenterOnScreen(r'.\Utils\clear_filter_button.png')
            pyautogui.doubleClick(x, y)
            break
        except Exception as e:
            print(e)
            pythonRPA.sleep(1)

    # Заполнение Табельного номера
    tab_nomer = pythonRPA.bySelector([{"title":"PRS_GR4","class_name":"TfrmFilterParams","backend":"win32"},{"ctrl_index":1},{"ctrl_index":4},{"ctrl_index":0}])
    tab_nomer.wait_appear(2)
    tab_nomer.click()
    pythonRPA.keyboard.write("002509")
    # Нажатие кнопки ОК
    ok  = pythonRPA.bySelector(
        [{"title": "PRS_GR4", "class_name": "TfrmFilterParams", "backend": "uia"}, {"ctrl_index": 2}, {"ctrl_index": 0},
         {"ctrl_index": 3}])
    ok.wait_appear(1)
    ok.click()

    # Выгрузка отчета
    for iterr in range(50):
        try:
            x, y = pyautogui.locateCenterOnScreen(r'.\Utils\report.png')
            pyautogui.doubleClick(x, y)
            break
        except Exception as e:
            print(e)
            pythonRPA.sleep(1)
    pythonRPA.sleep(5)

    for iterr in range(50):
        try:
            focus_on_report = pythonRPA.bySelector(
                [{"title": "Выбор отчета", "class_name": "TfrmRptLstRefer", "backend": "uia"}])
            focus_on_report.wait_appear(2)
            focus_on_report.maximize()
            x, y = pyautogui.locateCenterOnScreen(r'.\Utils\filter_of_report.png')
            pyautogui.doubleClick(x, y)
            break
        except Exception as e:
            print(e)
            pythonRPA.sleep(1)

    # Set focus on Filter app
    filter_s = pythonRPA.bySelector([{"title": "Фильтр", "class_name": "TfrmFilterParams", "backend": "uia"}])
    filter_s.wait_appear(15)
    filter_s.set_focus()

    filtering_by_name = pythonRPA.bySelector(
        [{"title": "Фильтр", "class_name": "TfrmFilterParams", "backend": "uia"}, {"ctrl_index": 0}, {"ctrl_index": 0},
         {"ctrl_index": 0}])
    filtering_by_name.click()
    pythonRPA.keyboard.write("увол")

    filtering_by_code = pythonRPA.bySelector(
        [{"title": "Фильтр", "class_name": "TfrmFilterParams", "backend": "uia"}, {"ctrl_index": 0}, {"ctrl_index": 1},
         {"ctrl_index": 0}])
    filtering_by_code.click()
    pythonRPA.keyboard.write('Z_160_PRIKAZI_UVOLENNIH')

    filter_ok = pythonRPA.bySelector(
        [{"title": "Фильтр", "class_name": "TfrmFilterParams", "backend": "uia"}, {"ctrl_index": 1}, {"ctrl_index": 0},
         {"ctrl_index": 2}])
    filter_ok.click()

    # openning the parametrs of the filter
    time.sleep(2)
    pythonRPA.keyboard.press('Enter')
    parametrs_of_filter = pythonRPA.bySelector(
        [{"title": "Параметры отчета ", "class_name": "TfrmRptPrmDialog", "backend": "uia"}])
    parametrs_of_filter.wait_appear(15)
    parametrs_of_filter.set_focus()
    pythonRPA.keyboard.press('Enter')

    parametr_filial = pythonRPA.bySelector(
        [{"title": "Параметры отчета ", "class_name": "TfrmRptPrmDialog", "backend": "uia"}, {"ctrl_index": 0},
         {"ctrl_index": 0}, {"ctrl_index": 0}])
    parametr_filial.click()
    pythonRPA.keyboard.write('00')
    pythonRPA.keyboard.press('Enter')
    # from_date = pythonRPA.bySelector([{"title":"Параметры отчета ","class_name":"TfrmRptPrmDialog","backend":"win32"},{"ctrl_index":0},{"ctrl_index":2},{"ctrl_index":0}])
    # to_date = pythonRPA.bySelector([{"title":"Параметры отчета ","class_name":"TfrmRptPrmDialog","backend":"win32"},{"ctrl_index":0},{"ctrl_index":4},{"ctrl_index":0}])

    today = date.today()
    today = today.strftime("%d%m%y")
    # from_date.click()
    pythonRPA.sleep(1)
    pythonRPA.keyboard.write("010120")
    pythonRPA.keyboard.press('Enter')
    # to_date.click()
    pythonRPA.sleep(1)
    pythonRPA.keyboard.write(today)
    parametr_ok = pythonRPA.bySelector(
        [{"title": "Параметры отчета ", "class_name": "TfrmRptPrmDialog", "backend": "uia"}, {"ctrl_index": 1},
         {"ctrl_index": 0}, {"ctrl_index": 1}])
    parametr_ok.click()
    pythonRPA.sleep(30)
    print("The report has been donwloaded sucsessfully !")

def Block_in_Colvir(names,nums,log_table):
    error = False
    while (1):
        try:
            # запуск Colvirtest (40TP)
            pythonRPA.application("C:\CBS_T/COLVIR.EXE").start()
            break
        except Exception as e:
            print('COLVIR CAN NOT BE STARTED   ' + e)

    unblocked_names = []
        # Выполнение входа
    log_in = pythonRPA.bySelector([{"title": "Вход в систему", "class_name": "TfrmLoginDlg", "backend": "uia"}])
    log_in.wait_appear(30)
    log_in.click()
    log_in.set_focus()
    pythonRPA.keyboard.write("colvir")
    pythonRPA.keyboard.press("Enter")
    pythonRPA.keyboard.write("test_2722")
    pythonRPA.keyboard.press('Enter')
    # ert_1908
    # konakayev_t
            # Перенаправление факуса на выбора режима
    mode_choosing = pythonRPA.bySelector(
        [{"title": "Выбор режима", "class_name": "TfrmCssMenu", "backend": "win32"}])
    mode_choosing.wait_appear(20)
    mode_choosing.set_focus()
    mode_s = pythonRPA.bySelector(
        [{"title": "Выбор режима", "class_name": "TfrmCssMenu", "backend": "win32"}, {"ctrl_index": 3},
         {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 0}])
    mode_s.wait_appear(5)
    mode_s.click()

    # Очистка фильтра задач
    pythonRPA.keyboard.press('SHIFT+CTRL+A')
    pythonRPA.keyboard.press('Delete')
    pythonRPA.keyboard.write('USRGRANT')
    pythonRPA.keyboard.press('Enter')

    # флаг того что мы должны скипнуть аккаунт и перейти на следющиего пользоватлея, изначально думаем то что этот аккаунт ункальный и проверяем его
    skip = False
    for name in names:
        try:
            if(not skip):
                while name!=names[0]:
                    try:
                        x,y = pyautogui.locateCenterOnScreen(r'.\Utils\filter_of_report.png')
                        pyautogui.click(x,y)
                        break
                    except:
                        continue

            # Перенаправление фокус экрана на фильтр
            search_filter = pythonRPA.bySelector([{"title": "Фильтр", "class_name": "TfrmFilterParams", "backend": "uia"}])
            search_filter.wait_appear(2)
            search_filter.set_focus()

            # Очистка фильтра поиска
            while 1:
                try:
                    x, y = pyautogui.locateCenterOnScreen(r'.\Utils\clear_filter_button.png')
                    pyautogui.doubleClick(x, y)
                    break
                except Exception as e:
                    print(e)

            # Заполнение наеменования
            name_line = pythonRPA.bySelector(
                [{"title": "Фильтр", "class_name": "TfrmFilterParams", "backend": "uia"}, {"ctrl_index": 1},
                 {"ctrl_index": 0},
                 {"ctrl_index": 0}])
            name_line.click()
            pythonRPA.keyboard.write(name)
            #Проверка на то что мы должны скипнутть или нет
            if(skip):
                pythonRPA.sleep(2)
            filter_ok = pythonRPA.bySelector([{"title": "Фильтр", "class_name": "TfrmFilterParams", "backend": "uia"},
                                              {"ctrl_index": 2}, {"ctrl_index": 0}, {"ctrl_index": 2}])
            if(skip):
                filter_ok.wait_appear(3)
            while 1:
                try:
                    filter_ok.click()
                    break
                except:
                    continue
            try:
                warning = pythonRPA.bySelector([{"title":"Подтверждение","class_name":"TMessageForm","backend":"win32"}])
                skip = warning.is_exists()
            except Exception as e:
                print(e)
            if (skip):
                pythonRPA.keyboard.press('right')
                pythonRPA.keyboard.press('Enter')
                log_table.append([name, "Colvir", "Не заблокирован", "Сотрудник не найден в системе", datetime.now().strftime("%H:%M:%S")])
                continue

            #Если мы не должны скипать то мы заходим на этот аккаунт и введем изминения
            else:
                # Вход на корректировку
                results_window = pythonRPA.bySelector(
                    [{"title": "Назначение профилей и полномочий позиций", "class_name": "TfrmUsrGrantList",
                      "backend": "uia"}])
                results_window.wait_appear(3)
                results_window.maximize()
                results_window.set_focus()
                code = "npos"
                while 1:
                    #Проверка на то что мы здесь выпервые или мы тут уже были
                    if(code!="npos"):
                        #Если мы тут впервые то нажимаем down
                        pythonRPA.keyboard.press('down')
                    pythonRPA.keyboard.press('Enter')
                    while 1:
                        try:
                            code_line = pythonRPA.bySelector([{"class_name":"TfrmUsrGrantGrpDetail","backend":"win32"},{"ctrl_index":0},
                                                      {"ctrl_index":0},{"ctrl_index":0},{"ctrl_index":10},{"ctrl_index":0}])
                            code_line.wait_appear(1)
                            break
                        except:
                         continue
                    if(code == code_line.texts()[0]):
                        # Мы заканчиваем эту иттерацию если мы впредыдущей итерации были здесь(Последняя учетная запись)
                        break
                    else:
                        code = code_line.texts()[0]

                    code = code_line.texts()[0]

                    # ФЛАГИ на проверку мы должны нажимать на эти кнопки или нет, изначально думаем что должны
                    archive_click = True
                    zapret_click = True

                    for i in range(2):
                        try:
                            # Здесь проверяем на галочку архива
                            x, y = pyautogui.locateCenterOnScreen(r'.\Utils\archive_true.png')
                            archive_click = False
                            break
                        except Exception as e:
                            print(e)

                    for i in range(2):
                        try:
                            x, y = pyautogui.locateCenterOnScreen(r'.\Utils\zapret_true.png')
                            # Здесь проверяем на галочку архива
                            zapret_click = False
                            break
                        except Exception as e:
                            print(e)
                    # Кнопка Пользователя
                    position = pythonRPA.bySelector([{ "class_name": "TfrmUsrGrantGrpDetail", "backend": "uia"}])
                    position1 = pythonRPA.bySelector([{"class_name":"TfrmUsrGrantGrpDetail","backend":"uia"},{"ctrl_index":4},{"ctrl_index":0}])
                    work_place = pythonRPA.bySelector(
                        [{"class_name": "TfrmUsrGrantGrpDetail", "backend": "uia"}, {"ctrl_index": 0},
                         {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 3}, {"ctrl_index": 0}])
                    change_work_place = bool(work_place.get_value()!='!!!!!!АРМ Пустой!' and work_place.is_exists())
                    if(archive_click or zapret_click or change_work_place):
                        # Нажатие на кнопку коррекции
                        # Перенаправление фокус экрана на окошку редактирование
                        position.wait_appear(3)
                        position.set_focus()
                        position1.click()
                        pythonRPA.keyboard.press('down',3)
                        pythonRPA.keyboard.press('enter')
                        if (change_work_place):
                            work_place.wait_appear(1)
                            work_place.click()
                            pythonRPA.keyboard.press('ctrl+A')
                            pythonRPA.keyboard.press('delete')
                            pythonRPA.keyboard.write("!!!!!!АРМ Пустой!")
                        # АРХИВ ЗАПРЕТ галочки и СОХРАНИТЬ
                        print(archive_click,zapret_click)
                        if(archive_click):
                            archive = pythonRPA.bySelector([{"class_name": "TfrmUsrGrantGrpDetail", "backend": "uia"},
                                                            {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 2}, {"ctrl_index": 5}])
                            archive.wait_appear(1)
                            archive.click()
                        if(zapret_click):
                            zapret = pythonRPA.bySelector([{"class_name": "TfrmUsrGrantGrpDetail",
                                                        "backend": "uia"}, {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 2},
                                                       {"ctrl_index": 4}])
                            zapret.wait_appear(1)
                            zapret.click()
                        save = pythonRPA.bySelector([{"class_name": "TfrmUsrGrantGrpDetail", "backend": "uia"}, {"ctrl_index": 4}, {"ctrl_index": 0}])
                        save.wait_appear(2)
                        save.click()
                        pythonRPA.keyboard.press('down',2)
                        pythonRPA.keyboard.press('enter')
                        pythonRPA.sleep(3)
                    # Закрытие Позиции пользователя(Окошка коррекции учетной записи)
                    position.close()

                    #Дальше переходим к следующему шагу (Настройка полномочии)
                    while 1:
                        try:
                            x, y = pyautogui.locateCenterOnScreen(r'.\Utils\polnom.png')
                            pyautogui.doubleClick(x, y)
                            break
                        except Exception as e:
                            print(e)


                    dependence_window = pythonRPA.bySelector([{"class_name":"TfrmHieFrmOrg","backend":"uia"}])
                    dependence_window.wait_appear(2)
                    dependence_window.set_focus()
                    dependence_window.maximize()
                    while(1):
                        try:
                            #нажимаем на кнопку двойной стрелки для того что бы мы голи нажатием пробелв удалять полномочии
                            x,y = pyautogui.locateCenterOnScreen(r'.\Utils\left_button.png')
                            pyautogui.click(x,y)
                            break
                        except:
                            continue

                    for i in range(2):
                        pythonRPA.keyboard.press('down')
                    pythonRPA.keyboard.press('right')
                    pythonRPA.keyboard.press('down')

                    pythonRPA.keyboard.press('space',20)
                    pythonRPA.sleep(2) # в бою изменить период задержки

                    pythonRPA.keyboard.press('up', 5)
                    for i in range(3):
                        pythonRPA.keyboard.press('down')
                    pythonRPA.keyboard.press('right')
                    pythonRPA.keyboard.press('down')

                    pythonRPA.keyboard.press('space', 20)
                    pythonRPA.sleep(2)  # в бою изменить период задержки

                    dependence_window.close()
                    pythonRPA.sleep(3)
                    warning = pythonRPA.bySelector([{"title":"Подтверждение","class_name":"TMessageForm","backend":"uia"}])
                    try:
                        if warning.is_exists():
                            pythonRPA.keyboard.press('Enter')
                            pythonRPA.sleep(1)
                            pythonRPA.keyboard.press('Enter')
                    except:
                        print('no changes made')
                    pythonRPA.sleep(3)
            # Если мы дошли до этой строки значит робот не сломался и заблокиравал аккаунта удачно !
            # Ввод в действие полномочии пользователей
            while 1:
                try:
                    x, y = pyautogui.locateCenterOnScreen(r'.\Utils\red.png')
                    pyautogui.doubleClick(x, y)
                    print('red button pressed')
                    break
                except Exception as e:
                    print(e)
            pythonRPA.sleep(3)
            pythonRPA.keyboard.press('Enter')
            pythonRPA.sleep(2)
            pythonRPA.keyboard.press('Enter')
            pythonRPA.sleep(2)
            log_table.append([name, "Colvir", "Заблокирован", "", datetime.now().strftime("%H:%M:%S")])
        except Exception as e:
            if not error:
             error = True
             Block_in_Colvir(names = [name],nums=nums,log_table=log_table)
             print(e)
            else:
                log_table.append([name, "Colvir", "Не заблокирован", "Техническая ошибка",datetime.now().strftime("%H:%M:%S")])
                print('bacccccc')
                continue
    #
    # for num in nums:
    try:
        os.system("taskkill /f /im COLVIR.EXE")
    except Exception as e:
        print(e)
    print('THE END')
