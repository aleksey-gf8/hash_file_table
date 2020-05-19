# -*- coding: utf-8-*-

import os
import sys
import glob
import time
import datetime
import hashlib
import openpyxl
from openpyxl.styles import Border, Side, Font, Alignment
from openpyxl import load_workbook
import docx


# types of files to be processed, skip the others
FILE_EXTENSION = {'.xlsx', '.xlsm', '.xls', '.docx', '.doc'}

# you need to skip this folder
NAME_SKIP_FOLDER = {'V:\\00 ', 'V:\\01 '}

# name of output file
OUT_FILE_NAME = r'Акт сдачи в архив электронных документов.xlsx'

# Руководитель Контрольной Службы
SKR_VALIDATOR = 'Атабиева М.И.'


def set_rightly_file_property(file, project_name, auditors_list):

    boss = auditors_list[0]
    if len(auditors_list) == 3:
        author = auditors_list[2]
    else:
        author = auditors_list[1]

    # check the current list item "file_list" this is a file?
    if not os.path.isfile(file):
        print(f'This is not file: {file}')
        return -20    

    find_dot = file.rfind('.')
    if find_dot == -1:
        print('i do not find dot.')
        return -21

    ext_file = file[find_dot:]
    if ext_file == '.xlsx':
        pass
        # set_xlsx_file_property(file, project_name, author, boss)
    elif ext_file == '.docx':
        set_docx_file_property(file, project_name, author, boss)
    elif ext_file == '.xlsm':
        pass


def set_xlsx_file_property(file_name, project_name, author, boss):
    wb = load_workbook(file_name)
    # блок Описание
    wb.properties.title = "ООО Группа Финансы" # Название
    wb.properties.subject = project_name # Тема, можно поставить имя папки проекта
    wb.properties.keywords = "Нет" # Теги, можно поставить "Нет"
    wb.properties.category ="Нет" # Категории, можно поставить "Нет"
    wb.properties.description = "Нет" # Комментарии, можно поставить "Нет"
    # блок Источник
    wb.properties.creator = author #Авторы, можно поставить "ФИО руководителя проверки"
    wb.properties.lastModifiedBy = boss # Кем сохранён, можно поставить "ФИО руководителя проверки"
    wb.properties.revision = "001" # Редакция, можно поставить "Нет"
    wb.properties.version = "002" # Номер версии, можно поставить "Нет"
    # блок Источник
    wb.properties.created=datetime.datetime(2019, 1, 31, 9, 0, 0) # Дата создания содержимого
    wb.properties.modified=datetime.datetime(2019, 2, 28, 9, 0, 0) # Дата последнего сохранения
    wb.properties.lastPrinted=datetime.datetime(2019, 3, 30, 9, 0, 0) # Последний вывод на печать
    # блок Содержание
    wb.properties.contentStatus="None" # Состояние содержимого
    wb.properties.language="RUS" # Язык
    # Визуально Никуда не идёт: wb.properties.identifier = "None"
    wb.save(file_name)


def set_docx_file_property(file_name, project_name, author, boss):
    doc = docx.Document(file_name)
    # блок Описание
    doc.core_properties.title = "ООО Группа Финансы" # Название, можно поставить имя папки "Авт сдачи в архив"
    doc.core_properties.subject = project_name # Тема, можно поставить имя папки проекта
    doc.core_properties.keywords = "Нет" # Теги, можно поставить "Нет"
    doc.core_properties.category = "Нет" # Категории, можно поставить "Нет"
    doc.core_properties.comments = "Нет" # Комментарии, можно поставить "Нет"
    # блок Источник
    doc.core_properties.author = author #Авторы, можно поставить "ФИО руководителя проверки"
    doc.core_properties.last_modified_by = boss # Кем сохранён, можно поставить "ФИО руководителя проверки"
    doc.core_properties.revision = 1 # Редакция, можно поставить "Нет"
    doc.core_properties.version= "002" # Номер версии, можно поставить "Нет"
    # блок Источник
    doc.core_properties.created = datetime.datetime(2020, 1, 31, 9, 0, 0) # Дата создания содержимого
    doc.core_properties.modified = datetime.datetime(2020, 2, 28, 9, 0, 0) # Дата последнего сохранения
    doc.core_properties.last_printed = datetime.datetime(2020, 3, 30, 9, 0, 0) # Последний вывод на печать
    # блок Содержание
    doc.core_properties.content_status ="None" # Состояние содержимого
    doc.core_properties.language='RUS' # Язык
    # Визуально Никуда не идёт: doc.core_properties.identifier='???'
    #debug-print(doc.core_properties)
    doc.save(file_name)
    time.sleep(1)


def get_auditors_list():
    """ returns a list of auditors who worked on the audit
        taken from a file ".\06 Аудит по существу\06.00 Содержание.xlsm"
        if found then return ['Boss_Name', 'Auditor_Name']
        if not found then return return ['НЕ найден', 'НЕ найден']
    """
    SOURCE_BOOKS = [
        "V:\\06 Аудит по существу\\06.00 Содержание.xlsm",
        "V:\\06  Аудит по существу\\06.00 Содержание.xlsm",
        "V:\\06 Аудит по существу\\06.00  Содержание.xlsm",
        "V:\\06  Аудит по существу\\06.00  Содержание.xlsm",
    ]
    FIND_STRING_BOOS = [
        'Руководитель задания:',
        'Руководитель проверки:',
    ]

    # checking the file's existence
    for source_book in SOURCE_BOOKS:
        if os.path.isfile(source_book):
            break
    else:
        print(f'This file does not exist: {source_book}')
        return ['НЕ найден', 'НЕ найден']
    # checking the existence of the sheet in the file
    wb = openpyxl.load_workbook(source_book)
    if wb.sheetnames.count('06') == 0:
        print(f'Sheet name 06 does not exist: {source_book}')
        return ['НЕ найден', 'НЕ найден']

    auditors_list = ['НЕ найден',]
    # читаем из excel-файла ФИО аудиторов
    sheet = wb['06']    # делаем лист '06' активным
    for row_count in range(1, 41):
        if sheet.cell(row=row_count, column=1).value in FIND_STRING_BOOS:
            auditors_list[0] = sheet.cell(row=row_count, column=2).value
            break

    for find_count in range(1, 41):
        if sheet.cell(row=find_count, column=1).value == 'Состав группы:':
            break
    else:
        auditors_list.append(auditors_list[0])
        return auditors_list

    for count in range(find_count+1, find_count+5):        
        if sheet.cell(row=count, column=1).value == 1:
            auditors_list.append(sheet.cell(row=count, column=2).value)
        elif sheet.cell(row=count, column=1).value == 2:
            auditors_list.append(sheet.cell(row=count, column=2).value)
        elif sheet.cell(row=count, column=1).value == 3:
            auditors_list.append(sheet.cell(row=count, column=2).value)
    
    if len(auditors_list) == 1:
        auditors_list.append(auditors_list[0])

    return auditors_list            


def get_source_path():
    """ returns the path to the folder where the data is stored (files)
    """
    if "MY_SOURCE_PATH" in os.environ:
        # имя каталога должно быть без кавычек
        return os.environ["MY_SOURCE_PATH"].replace('"', '')
    else:
        return os.getcwd()


def get_result_path():
    """ returns the path to the folder where you want to save data (files)
    
    if True:
        # save to the current folder
        result_path = os.getcwd()
    else:
        # save to the project folder
        result_path = source_path
    return result_path
    """
    result_path = os.getcwd()
    return result_path

def file_is_needed(file):
    """ Check - this file needs to be processed. If need, then return True,
    otherwise - False.
    """
    find_dot = file.rfind('.')
    if find_dot == -1:
        return False
    elif (file[find_dot:] in FILE_EXTENSION):
        find_temp = file.rfind('\\~$')
        if find_temp == -1:
            return True
        else:
            print(f'This is temporary file={file}, skip')
            return False
    else:
        return False

def skip_this_folder(name):
    """ Check - you need to skip this folder. If need, then return True,
    otherwise - False.
    """
    return (name[0:6] in NAME_SKIP_FOLDER)


def get_hash_md5(filename):
    """ get hash-MD5 of filename object
    """
    with open(filename, 'rb') as f:
        m = hashlib.md5()
        while True:
            data = f.read(8192)
            if not data:
                break
            m.update(data)
        return m.hexdigest()


def remove_link_to_v():
    """ Remove the Symbolic link to disk V:
    """
    if os.path.exists("V:\\"):
        subst_command = 'subst /d V:'
        return os.system(subst_command)


def create_link_to_v(source_path):
    """ get a list of all items in the current directory
    """
    # Удаляем Символическую ссылку на диск V: если она существует
    remove_link_to_v()
    # Подключаем Симвлолическую ссылку диск V: указывает на "source_path"
    subst_command = 'subst V: "{0}"'.format(source_path)
    ret_code = os.system(subst_command)
    return ret_code


def create_table_and_set_file_property(
        file_list,
        project_name,
        auditors_list):
    """ create а table from a 'file_list' each row consists of:
    file name + file hash + author + validator
    """

    validator = auditors_list[0]    # Проверил
    if len(auditors_list) == 3:
        author = auditors_list[2]   # Автор
    else:
        author = auditors_list[1]   # Автор

    result_table = []
    for file in file_list:
# check the current list item "file_list" - this should be processed or skipped?
        if not os.path.isfile(file):
            print(f'We need to skip because it is a folder: {file}')
            continue
        if skip_this_folder(file):
            print(f'We need to skip files in this folder: {file}')
            continue
        if not file_is_needed(file):
            print(f'We need to skip because this file is not needed: {file}')
            continue

        print(f'обрабатываем файл: {file}')
        set_rightly_file_property(file, project_name, auditors_list)
        # create new record in the result_table
        curr_line = {}
        curr_line.update({'file-name': file[2:]})
        curr_line.update({'Hash-MD5': get_hash_md5(file)})
        curr_line['author'] = author
        curr_line['validator'] = validator
        result_table.append(curr_line)
    return result_table


def save_table_to_xlsx_file(project_name, result_table, result_path):
    """ write 'result_table' to xlsx file """

    wb = openpyxl.Workbook()
    wb.create_sheet(title='Список файлов', index=0)
    sheet = wb['Список файлов']
    # create a header
    col = 1
    row = 1
    sheet.merge_cells('A1:E1')
    sheet['A1'] = 'АКТ'
    sheet['A1'].alignment = Alignment(horizontal="center", vertical="center")
    sheet['A1'].font = Font(bold=True)
    sheet.merge_cells('A2:E2')
    sheet['A2'] = 'сдачи в архив электронных документов по проекту ' + project_name
    sheet['A2'].alignment = Alignment(horizontal="center", vertical="center")
    sheet['A2'].font = Font(bold=True)
    # create a table-header row
    bbb = Side(border_style="medium", color="000000")
    sheet['A4'] = '№ п/п'
    sheet.column_dimensions['A'].width = 7
    sheet['A4'].font = Font(bold=True)
    sheet['A4'].border = Border(top=bbb, left=bbb, right=bbb, bottom=bbb)

    sheet['B4'] = 'Имя файла'
    sheet.column_dimensions['B'].width = 100
    sheet['B4'].font = Font(bold=True)
    sheet['B4'].border = Border(top=bbb, left=bbb, right=bbb, bottom=bbb)
    # sheet['C1'] = 'дата последнего изменения'
    sheet['C4'] = 'Хэш-сумма (MD5) файла'
    sheet.column_dimensions['C'].width = 37
    sheet['C4'].font = Font(bold=True)
    sheet['C4'].border = Border(top=bbb, left=bbb, right=bbb, bottom=bbb)

    sheet['D4'] = 'Файл создал'
    sheet.column_dimensions['D'].width = 20
    sheet['D4'].font = Font(bold=True)
    sheet['D4'].border = Border(top=bbb, left=bbb, right=bbb, bottom=bbb)

    sheet['E4'] = 'Файл проверил'
    sheet.column_dimensions['E'].width = 20
    sheet['E4'].font = Font(bold=True)
    sheet['E4'].border = Border(top=bbb, left=bbb, right=bbb, bottom=bbb)

    row = 5
    thin = Side(border_style="thin", color="000000")
    for line_dict in result_table:
        row += 1
        col = 1
        cell = sheet.cell(row=row, column=col)
        # this is a row counter, you need to take into account the empty rows
        # in the table header
        cell.value = row-5
        cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

        for __, val in line_dict.items():
            col += 1
            cell = sheet.cell(row=row, column=col)
            cell.value = val
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

    col = 1
    cell = sheet.cell(row=row+4, column=col)
    cell.value = f'Файлы принял на хранение в архив'
    cell = sheet.cell(row=row+5, column=col)
    cell.value = f'Руководитель Контрольной Службы:'
    cell = sheet.cell(row=row+5, column=col+3)
    cell.value = f'{SKR_VALIDATOR}'
    
    file_name = result_path + '\\' + project_name + '_' + OUT_FILE_NAME
    wb.save(file_name)


def process_create_file_table(source_path, result_path):
    """ Creates a list of all files as: file name, hash, author """

    ret_code = create_link_to_v(source_path)
    if ret_code != 0:
        print(f'Символ.ссылка V: НЕ СДЕЛАНА! Код возврата subst: {ret_code}')
        return 10
    print(f'Символ.ссылка СДЕЛАНА V: -> {source_path}')
    try:
        # создаём рекурсивно список объектов внутри "V:"
        file_list = glob.glob('V:\\**\\*', recursive=True)
        if file_list is None:
            print('A list of all items in current directory is not created')
            # list of files not created
            return 20

        auditors_list = get_auditors_list()
        print('auditors_list = ', *auditors_list)
        project_name = source_path[source_path.rindex('\\') + 1:]
        print(f'project_name = {project_name}')

        result_table = create_table_and_set_file_property(
            file_list,
            project_name,
            auditors_list
            )
        save_table_to_xlsx_file(project_name, result_table, result_path)
    finally:
        remove_link_to_v()
    return 0


def main(argv=None):

    source_path = get_source_path()
    print(f"source_path = {source_path}")
    result_path = get_result_path()
    print (f"result_path = {result_path}")

    ret_code = process_create_file_table(
        source_path,
        result_path
    )

    print(f"return code = {ret_code}")
    return ret_code


if __name__ == '__main__':
    sys.exit(main(sys.argv))