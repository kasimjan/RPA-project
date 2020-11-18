import pandas as pd
from Sources.documentolog_functions import start_documentolog, stop_documentolog, block_documentolog, reload_driver
from Sources.other_functions import write_logs_into_file_2

log_table = [["ФИО", "Статус блокировки", "Комментарии"]]

df = pd.read_excel(r"C:\Users\mazhit.e\Desktop\list_to_block_doc.xlsx") # can also index sheet by name or fetch all sheets
mylist = df['Names'].tolist()
print(len(mylist))
only_strings = [x for x in mylist if isinstance(x, str)]
print(len(only_strings))

for name in only_strings:
    name = name.split("(")[0].strip()
    print(name)
    try:
        start_documentolog()
        log_documentolog_status, log_comment2 = block_documentolog(name)
        log_table.append([name, log_documentolog_status, log_comment2])
        reload_driver()
    except:
         log_table.append([name, "Не заблокирован", "Техническая ошибка"])
stop_documentolog()

write_logs_into_file_2(log_table)
