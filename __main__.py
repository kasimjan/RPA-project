import wmi
from pyPythonRPA.Robot import pythonRPA
from Sources.bpm_selenium_ie import block_bpm, start_bpm, stop_bpm
from Sources.documentolog_functions import block_documentolog, start_documentolog, stop_documentolog, reload_driver
from Sources.other_functions import write_logs_into_file, Close_Excel, telegram_bot_sendtext
from Sources.report import delete_xml_files, get_path, get_names,get_nums
from Sources.Colvir import download_report, Block_in_Colvir
from Sources.AD import block as AD_block
from Sources.ColvirAdmin import ColvirAdmin_block
from time import sleep
import os
from datetime import datetime
from datetime import date

def main_function():
    # delete_xml_files(r'C:\Users\mazhit.e\AppData\Local\Temp)
    # download_report()

    # try:
    #     os.system("taskkill /f /im COLVIR.EXE")
    # except Exception as e:
    #     print(e)
    # Close_Excel()

    log_table = [["ФИО", "Система", "Статус", "Комментарии", "Время блокировки"]]
    names = get_names(r'C:\Users\mazhit.e\Desktop\97.xml') #getpath
    # names = ['Адилова Айгуль Сарсенбаевна']
    nums = get_nums(r'C:\Users\mazhit.e\Desktop\97.xml')
    if len(names):
        ## Блокировка в системе Active directory
        try:
            AD_block(names, log_table)
        except Exception as e:
            print(e)


        # #Блокировка в системе КОЛВИР АДМИН
        try:
            ColvirAdmin_block(names, log_table)
        except Exception as e:
            print(e)
        try:
            os.system("taskkill /f /im CBSADM.EXE")
        except Exception as e:
            print(e)


        #Блокировка в системе COLVIR.exe
        try:
            Block_in_Colvir(names,nums, log_table)
        except Exception as e:
            print(e)

        print("-------- Documentolog Started -------")
        # Блокировка на сайте Documentolog
        for name in names:
            try:
                start_documentolog()
                log_documentolog_status, log_comment2 = block_documentolog(name)
                log_table.append([name, "Documentolog", log_documentolog_status, log_comment2, datetime.now().strftime("%H:%M:%S")])
                reload_driver()
            except:
                log_table.append([name, "Documentolog", "Не заблокирован", "Техническая ошибка", datetime.now().strftime("%H:%M:%S")])
        stop_documentolog()
        print("-------- Documentolog Ended -------")

        print("-------- BPM Started -------")
        # Блокировка на сайте BPM
        start_bpm()
        for name in names:
            try:
                log_bpm_status, log_comment = block_bpm(name)
                log_table.append([name, "BPM", log_bpm_status, log_comment, datetime.now().strftime("%H:%M:%S")])
            except Exception as e:
                print(e)
                try:
                    stop_bpm()
                    start_bpm()
                    log_bpm_status, log_comment = block_bpm(name)
                    log_table.append([name, "BPM", log_bpm_status, log_comment, datetime.now().strftime("%H:%M:%S")])
                except:
                    log_table.append([name, "BPM", "Не заблокирован", "Техническая ошибка", datetime.now().strftime("%H:%M:%S")])
        stop_bpm()
        print("-------- BPM Ended -------")

        #Логирование
        sleep(3)
        print(log_table)
        write_logs_into_file(log_table)
        telegram_bot_sendtext("Добрый вечер! Блокировка сотрудников за "+date.today().strftime("%d.%m.%y")+" выполнена. Статусы можете посмотреть в файле, отправленном ниже.", r".\Files\Logs\logs "+date.today().strftime("%d.%m.%y")+".xlsx")

finished = False
while finished == False:
    try:
        main_function()
        finished = True
    except:
        try:
            f = wmi.WMI()
            for p in f.Win32_Process():
                if p.name == 'mmc.exe':
                    p.Terminate()
        except Exception as e:
            print(e)
        os.system("taskkill /f /im COLVIR.EXE")
        os.system("taskkill /f /im CBSADM.EXE")
        stop_bpm()
        stop_documentolog()
        main_function()
        finished = True

