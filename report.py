import os.path
from os import path
import xmltodict
import os
#the following function will convert xml report file to the jason file, then will return the list with names of employees who needed to block

#
#1 Fisrtly delete all xml files before get report
#2 get names
#get_names(get_path())

def get_names(path):
    with open(path, encoding='utf-8') as xml_file:
        data_dict = dict(xmltodict.parse(xml_file.read()))
        xml_file.close()
    i = 4
    names = []
    while 1:
        try:
            name = (data_dict['Workbook']['Worksheet']['Table']['Row'][i]['Cell'][2]['Data']['#text'])
            names.append(name)
            print(str(i-3)+")"+name)
            i += 1
        except Exception as e:
            break
    return names

def get_nums(path):
    with open(path, encoding='utf-8') as xml_file:
        data_dict = dict(xmltodict.parse(xml_file.read()))
        xml_file.close()
    i = 4
    nums = []
    while 1:
        try:
            num = (data_dict['Workbook']['Worksheet']['Table']['Row'][i]['Cell'][1]['Data']['#text'])
            nums.append(num)
            print(str(i-3)+")"+num)
            i += 1
        except Exception as e:
            break
    return nums



get_nums(r'C:\Users\mazhit.e\Desktop\97.xml')

#the following functoin will delete the xml files after using them
def delete_xml_files(abs_path = r'C:\Users\mazhit.e\AppData\Local\Temp'):
    try:
        files_and_directories = os.listdir(abs_path)
        for file_or_directory in files_and_directories:
            if (file_or_directory.find('.') == -1):
                directory = os.listdir(r'C:/Users/mazhit.e/AppData/Local/Temp/' + file_or_directory)
                # print(r'C:/Users/mazhit.e/AppData/Local/Temp/' + file_or_directory)
                for file in directory:
                    # print('   ', file)
                    if (file[-4:]) == '.xml':
                        # C:\Users\mazhit.e\AppData\Local\Temp\2\627.xml
                        abs_path = 'C:/Users/mazhit.e/AppData/Local/Temp/' + file_or_directory + '/' + file
                        os.remove(abs_path)
                        print(abs_path, " removed well!")
            elif (file_or_directory[-4:]) == '.xml':
                abs_path = 'C:/Users/mazhit.e/AppData/Local/Temp/' + file_or_directory
                if path.exists(abs_path):
                    os.remove(abs_path)
                    print(abs_path, " removed well!")
    except Exception as e:
        print(e)

#the following function will return the absalute path of the report (xml file)

def get_path():
    files_and_directories = os.listdir(r'C:\Users\mazhit.e\AppData\Local\Temp')
    for file_or_directory in files_and_directories:
        if(file_or_directory.find('.')==-1):
            directory = os.listdir(r'C:/Users/mazhit.e/AppData/Local/Temp/' +file_or_directory)
            print(r'C:/Users/mazhit.e/AppData/Local/Temp/' +file_or_directory)
            for file in directory:
                print('   ', file)
                if (file[-4:]) == '.xml':
                    # C:\Users\mazhit.e\AppData\Local\Temp\2\627.xml
                    abs_path = 'C:/Users/mazhit.e/AppData/Local/Temp/' +file_or_directory+'/'+file
                    print(" Absolute path found well!", abs_path)
                    return abs_path
        elif (file_or_directory[-4:]) == '.xml':
            abs_path = 'C:/Users/mazhit.e/AppData/Local/Temp/' + file_or_directory
            print(abs_path)
            if path.exists(abs_path):
                print(" Absolute path found well!", abs_path)
                return abs_path