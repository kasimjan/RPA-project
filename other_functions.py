import xlsxwriter
import os
from datetime import date
import webbrowser
import subprocess
from pyPythonRPA.Robot import pythonRPA
from time import sleep
import requests
import telebot

def telegram_bot_sendtext(bot_message, filename):
    bot_token = '1453385041:AAFfMNJaaviK6w0JV_mCxpSe5U2G0EpasQY'
    bot_chatID = '-1001462607640'
    tb = telebot.TeleBot(bot_token)
    doc = open(filename, 'rb')
    tb.send_message(bot_chatID, bot_message)
    tb.send_document(bot_chatID, doc)

def write_logs_into_file(log_table):
    workboook = xlsxwriter.Workbook(r".\Files\Logs\logs "+date.today().strftime("%d.%m.%y")+".xlsx")
    worksheet = workboook.add_worksheet()

    for i in range (len(log_table[0])):
        for j in range(len(log_table)):
            worksheet.write(j, i, log_table[j][i])

    workboook.close()

def Close_Excel():

    for i in range(4):
        try:
            os.system("taskkill /f /im EXCEL.EXE")
        except Exception as e:
         print(e)
         break

def check_ie_driver_settings():
    ie_subprocess = subprocess.Popen(r'"C:\Program Files\Internet Explorer\iexplore.exe" www.google.com')
    settings_button = pythonRPA.bySelector([{"title": "Google - Internet Explorer", "class_name": "IEFrame", "backend": "uia"}, {"depth_start": 4, "depth_end": 4, "title": "Tools", "control_type": "Button"}])
    settings_button.wait_appear(5)
    settings_button.click()
    sleep(0.2)
    pythonRPA.keyboard.press("down", 10)
    pythonRPA.keyboard.press("enter")
    sleep(0.2)
    pythonRPA.keyboard.press("ctrl+tab")

    enabled_checkbox = pythonRPA.byImage("./Utils/enable_protected_mode_ie_2.png")
    if enabled_checkbox.is_exists():
        enabled_checkbox.click()

    sleep(0.2)
    pythonRPA.keyboard.press("enter")
    sleep(0.2)
    pythonRPA.keyboard.press("enter")

    os.system("taskkill /f /im iexplore.exe")
    # ie = webbrowser.get(webbrowser.iexplore)
    # ie.open('google.com')

#check_ie_driver_settings()

def write_logs_into_file_2(log_table):
    workboook = xlsxwriter.Workbook(r"..\Files\Logs\documentolog_logs "+date.today().strftime("%d.%m.%y")+".xlsx")
    worksheet = workboook.add_worksheet()

    for i in range (len(log_table[0])):
        for j in range(len(log_table)):
            worksheet.write(j, i, log_table[j][i])

    workboook.close()