import zipfile
from email.header import make_header, decode_header
import email.message
import csv
import json
from base64 import b64decode
import os
from sys import argv
from itertools import count

save_att_enable = True  # включить сохранение аттачментов (отладочная опция)
eml_ext = '.eml'
content_exclude = ('multipart/alternative',
                   'text/plain',
                   'text/calendar',
                   'multipart/mixed',
                   'text/html',
                   'multipart/related',
                   'multipart/report',
                   'message/delivery-status',
                   'message/rfc822'
                   )

file_name_exclude = ('None', 'image001.png')

c_type_list = []

# file_name = 'd:\\tmp\\test_mbox.zip'
file_name = ''
try:
    script_name, f_name = argv
except ValueError:
    print("Параметр командной строки не найден.")
    f_name = input("Введите имя файла: ")
else:
    print("Имя файла передано из командной строки.")
finally:
    if os.path.exists(f_name):
        file_name = f_name
    else:
        print(f"Файл {f_name} не найден. Выход.")
        exit(1)

file_path = os.path.dirname(file_name)
export_path = file_name + '_export\\'
if not os.path.exists(export_path):
    print(f"Создание каталога экспорта: {export_path}")
    os.mkdir(export_path)

report_csv = f"{export_path}{os.path.basename(file_name)}_report.csv"
report_json = f"{export_path}{os.path.basename(file_name)}_report.json"

eml_info_list = [] # список словарей для сбора инфы по емайлам

print(f"Открываем архив {file_name}...")
with zipfile.ZipFile(file_name) as my_zip:
    names: list[str] = my_zip.namelist()
    print(f"Найдено {len(names)} элементов в архиве.")
    # print(names) # debug

    # перебираем элементы архива
    for zip_item in names:
        if zip_item.find(eml_ext) != -1:
            print(f"\nОткрываем файл сообщения {zip_item} :")

            # начинаем перебирать элементы архива
            with my_zip.open(zip_item, 'r') as file_item:
                msg = email.message_from_bytes(file_item.read())
                msg_properties_dict = {"Filename":zip_item}
                print(f"From: {msg.get('From')}")
                print(f"To: {msg.get('To')}")
                print(f"Date: {msg.get('Date')}")

                # сбор инфы по атрибутам письма
                msg_properties_dict["From"] = msg.get('From')
                msg_properties_dict["To"] = msg.get('To')
                msg_properties_dict['Date'] = msg.get('Date')
                msg_properties_dict["Subject"] = ''
                msg_properties_dict['Attachments'] = ''

                try:
                    msg_properties_dict["Subject"] = str(make_header(decode_header(msg.get('Subject'))))
                except TypeError:
                    msg_properties_dict["Subject"] = f"(DECODE ERROR!!!) {msg.get('Subject')}"

                if msg.is_multipart():
                    file_name_list = []
                    for part in msg.walk():  # проходим по частям сообщения
                        c_t_name = part.get_content_type()
                        if c_t_name not in content_exclude:
                            if c_t_name not in c_type_list:
                                c_type_list.append(c_t_name)
                            p_f_name = part.get_filename()
                            file_name_list.append(p_f_name)
                            p_file_data = b64decode(part.get_payload())
                            if type(p_f_name) is not None:
                                try:
                                    ss = str(make_header(decode_header(p_f_name)))
                                except TypeError:
                                    ss = f" DECODE ERR: {p_f_name}"
                                att_full_name = export_path + ss
                                # ищем не существует ли уже такой файл.
                                # Если существует, подставляем другое имя со счетчиком в начале
                                if os.path.exists(att_full_name):
                                    for ii in count(1):
                                        reserve_name = f"{export_path}{str(ii)}_{ss}"
                                        if not os.path.exists(reserve_name):
                                            print(f"файл существует, резервное имя: {reserve_name}")
                                            att_full_name = reserve_name
                                            break
                                        elif ii > 10000:
                                            break
                                # сохраняем файл
                                try:
                                    if save_att_enable:
                                        with open(att_full_name, "wb") as att:
                                            att.write(p_file_data)
                                except OSError:
                                    print(f"Ошибка записи файла {ss}")
                                else:
                                    print(f"Сохранен файл: {ss}")
                    msg_properties_dict['Attachments'] = str(file_name_list)
                else: # message is not multipart
                    print("Сообщение не содержит частей (вложений).")
                eml_info_list.append(msg_properties_dict)

#print(f"итого имен файлов: {len(file_name_list)}")
print(f"итого типов контента: {len(c_type_list)}")
print(f"Content types: {c_type_list}")

# сохраняем структуру отчета в CSV файле
with open(report_csv, "w", newline='') as csvfile:
    csvwriter = csv.DictWriter(csvfile, eml_info_list[0].keys())
    csvwriter.writeheader()
    for el in eml_info_list:
        try:
            csvwriter.writerow(el)
            #print(el)
        except UnicodeEncodeError:
            print(f"{el['Filename']}: UnicodeEncodeError: 'charmap' codec can't encode character")

# сохраняем структуру отчета в JSON
with open(report_json, 'w') as jsonfile:
    json.dump(eml_info_list, jsonfile)
