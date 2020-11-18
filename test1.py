from datetime import datetime
import os.path
from os import path
import os
from Sources import report
import keyboard
from pyautogui import *
import time
from pyPythonRPA.Robot import pythonRPA
from Sources.other_functions import telegram_bot_sendtext
from datetime import date

# # def getpath():
# #     files_and_directories = os.listdir(r'C:\Users\mazhit.e\AppData\Loc5al\Temp')
# #     for file_or_directory in files_and_directories:
# #         if(file_or_directory.find('.')==-1):
# #             directory = os.listdir(r'C:/Users/mazhit.e/AppData/Local/Temp/' +file_or_directory)
# #             print(r'C:/Users/mazhit.e/AppData/Local/Temp/' +file_or_directory)
# #             for file in directory:
# #                 print('   ', file)
# #                 if (file[-4:]) == '.xml':
#                     # C:\Users\mazhit.e\AppData\Local\Temp\2\627.xml
#                     abs_path = 'C:/Users/mazhit.e/AppData/Local/Temp/' +file_or_directory+'/'+file
#                     print(" Absolute path found well!", abs_path)
#                     return abs_path
#         elif (file_or_directory[-4:]) == '.xml':
#             abs_path = 'C:/Users/mazhit.e/AppData/Local/Temp/' + file_or_directory
#             print(abs_path)
#             if path.exists(abs_path):
#                 print(" Absolute path found well!", abs_path)
#                 return abs_path
#         # else: print(file_or_directory)
# 
# report.delete_xml_files()
# log_table = [['ФИО', 'Система', 'Статус блокировки', 'Комментарии', 'Время блокировки'],
#              ['test test test', 'Active Directory', 'Заблокирован', '', '13/11/2020 17:22:44'],
#              ['test test test', 'CBSADM', 'Не заблокирован', 'Сотрудник не найден в системе', '13/11/2020 17:23:13'],
#              ['Абдраимов Канат Турсынбаевич', 'CBSADM', 'Не заблокирован', 'Техническая ошибка', '13/11/2020 17:23:15'],
#              ['Кантемирова Надежда Викторовна', 'CBSADM', 'Заблокирован', '', '13/11/2020 17:23:48'],
#              ['Абдраймов Рахимжан Амирбекович', 'CBSADM', 'Не заблокирован', 'Сотрудник не найден в системе', '13/11/2020 17:23:55'],
#              ['test test test', 'CBSADM', 'Не заблокирован', 'Сотрудник не найден в системе', '13/11/2020 17:24:06'],
#              ['Абдраимов Канат Турсынбаевич', 'CBSADM', 'Заблокирован', '', '13/11/2020 17:24:26'],
#              ['Кантемирова Надежда Викторовна', 'CBSADM', 'Заблокирован', '', '13/11/2020 17:24:40'],
#              ['Абдраймов Рахимжан Амирбекович', 'CBSADM', 'Не заблокирован', 'Сотрудник не найден в системе', '13/11/2020 17:24:47'],
#              ['test test test', 'Colvir', 'Не заблокирован', 'Сотрудник не найден в системе', '13/11/2020 17:25:04'],
#              ['Абдраимов Канат Турсынбаевич', 'Colvir', 'Заблокирован', '', '13/11/2020 17:25:56'],
#              ['Кантемирова Надежда Викторовна', 'Colvir', 'Заблокирован', '', '13/11/2020 17:28:45'],
#              ['Абдраймов Рахимжан Амирбекович', 'Colvir', 'Не заблокирован', 'Сотрудник не найден в системе', '13/11/2020 17:28:49'],
#              ['test test test', 'Documentolog', 'Не заблокирован', 'Сотрудник не найден в системе', '13/11/2020 17:29:00'],
#              ['Абдраимов Канат Турсынбаевич', 'Documentolog', 'Заблокирован', '', '13/11/2020 17:29:26'],
#              ['Кантемирова Надежда Викторовна', 'Documentolog', 'Заблокирован', '', '13/11/2020 17:29:49'],
#              ['Абдраймов Рахимжан Амирбекович', 'Documentolog', 'Заблокирован', '', '13/11/2020 17:30:13'],
#              ['test test test', 'BPM', 'Не заблокирован', 'Сотрудник не найден в системе', '13/11/2020 17:30:33'],
#              ['Абдраимов Канат Турсынбаевич', 'BPM', 'Заблокирован', '', '13/11/2020 17:30:57'],
#              ['Кантемирова Надежда Викторовна', 'BPM', 'Не заблокирован', 'Сотрудник не найден в системе', '13/11/2020 17:31:02'],
#              ['Абдраймов Рахимжан Амирбекович', 'BPM', 'Заблокирован', '', '13/11/2020 17:31:25']]
# for row in log_table:
#     if row[1] == 'CBSADM':
#         log_table.remove(row)
# from Sources.other_functions import write_logs_into_file
# # write_logs_into_file(log_table)
# for row in log_table:
#     print(row)
def change_row(log_table,row1,name,system,status,comment,time):
    for row in log_table:
        if row==row1:
            row[0] = name
            row[1] = system
            row[2] = status
            row[3] = comment
            row[4] = time
names = [['Адилова Айгуль Сарсенбаевна']]
# telegram_bot_sendtext("Добрый вечер! Блокировка сотрудников за "+date.today().strftime("%d.%m.%y")+" выполнена. Статусы можете посмотреть в файле отправленном ниже файле.", r".\Files\Logs\logs "+date.today().strftime("%d.%m.%y")+".xlsx")
for name in names:
    log_table = [["ФИО", "Система", "Статус блокировки", "Комментарии", "Время блокировки"], ['Адилова Айгуль Сарсенбаевна', 'CBSADM', 'Не Заблокирован', '', '14:39:04']]
    for row in log_table:
        change_row(log_table,row,'Адилова Айгуль Сарсенбаевна', 'CBSADM', 'Заблокирован', '', '14:30')
        # if row[1] == 'CBSADM':
        #     row[0] = name
        #     row[1] = "CBSADM"
        #     row[2] = "Заблокирован"
        #     row[3] = ""
        #     row[4] = datetime.now().strftime("%H:%M:%S")
print(log_table)

