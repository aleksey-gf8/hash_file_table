import os
import time
import glob
import hashlib
import openpyxl


# types of files to be processed
file_extension = {'.xlsx', '.xls', '.docx', '.doc'}

# skip_this_folder
name_skip_folder = {'V:\\00 ', 'V:\\01 '}

# global var
auditors_list = []

OUT_FILE_NAME = r'Акт сдачи в архив электронных документов'
SKR_VALIDATOR = 'Иванов И.И.'


def get_auditors_list(source_path):
    """ returns a list of auditors who worked on the audit
        taken from a file ".\05  Аудит  СВК\05.00 Содержание.xlsx"
    """
    source_books = ["V:\\05  Аудит  СВК\\05.00 Содержание.xlsx",
                    "V:\\05 Аудит СВК\\05.00 Содержание.xlsx",
                    "V:\\05 Аудит  СВК\\05.00 Содержание.xlsx",
                    "V:\\05  Аудит СВК\\05.00 Содержание.xlsx",
    ]
    not_found = True
    for source_book in source_books:
        if not os.path.isfile(source_book):
            continue
        else:
            not_found = False
            break
        
    if not_found:     
        print('"This file does not exist": {0}'.format(source_book))
        return None  
    # читаем excel-файл
    wb = openpyxl.load_workbook(source_book)
    if wb.sheetnames.count('12') == 0:
        print('"Sheet name 12 does not exist": {0}'.format(source_book))
        return None
    result = []
    # делаем 12 лист активным
    wb.active = 11
    # получаем активный лист
    sheet = wb.active
    if sheet['A29'].value == 1:
        result.append(sheet['B29'].value)
    if sheet['A31'].value == 'Руководитель проверки:':
        result.append(sheet['B31'].value)
    else:
        numbers_by_row = [32, 33, 34, 35]
        for row_count in numbers_by_row:
            if sheet.cell(row = row_count, column = 1).value == 'Руководитель проверки:':
                result.append(sheet.cell(row = row_count, column = 2).value)
                break
    print('result= ', result)
    return result


def get_source_path():
    """ returns the path to the folder where the data is stored (files)
    """
    if "MY_SOURCE_PATH" in os.environ:
        # имя каталога без кавычек
        return os.environ["MY_SOURCE_PATH"].replace('"', '')
    else:
        return os.getcwd()


def file_is_needed(file):
    """ Check - this file needs to be processed. If need, then return True,
    otherwise - False.
    """
    find_dot = file.rfind('.')
    if find_dot == -1:
        return False
    return (file[find_dot:] in file_extension)


def skip_this_folder(file):
    """ Check - you need to skip this folder. If need, then return True,
    otherwise - False.
    """
    return (file[0:6] in name_skip_folder)


def remove_link_to_v():
    """ Remove the Symbolic link to disk V: """
    if os.path.exists("V:\\"):
        subst_command = 'subst /d V:'
        return os.system(subst_command)


def get_hash_md5(filename):
    """ get hash-MD5 of filename object """

    with open(filename, 'rb') as f:
        m = hashlib.md5()
        while True:
            data = f.read(8192)
            if not data:
                break
            m.update(data)
        return m.hexdigest()


def get_file_list(source_path):
    """ get a list of all items in the current directory
    """

    # Удаляем Символическую ссылку на диск V: если она существует
    remove_link_to_v()

    # Подключаем Симвлолическую ссылку диск V: указывает на "source_path"
    subst_command = 'subst V: "{0}"'.format(source_path)
    print(subst_command)
    ret_code = os.system(subst_command)

    if ret_code != 0:
        print('Символ.ссылка НЕ СДЕЛАНА! V:-> {0} Код возврата subst: {1}'.
              format(source_path, ret_code))
        # завершаем и переходим к следующей итерации цикла по source_dir.
        return None
    else:
        print('Символ.ссылка СДЕЛАНА диск V: -> {0} Код возврата subst: {1}'.
              format(source_path, ret_code))

    # создаём рекурсивно список объектов внутри "V:\ (src_dir_level1)"
    return glob.glob('V:\\**\\*', recursive=True)


def create_table(file_list):
    """ create а table from a 'file_list' each row consists of:
    file name + file hash + author + validator
    """
    global auditors_list
    auditors_list = get_auditors_list(source_path)
    if auditors_list:
        author = auditors_list[0]
        validator = auditors_list[1]
    else:
        print("no read author")
        author = 'author'
        validator = 'validator'        
    result_table = []
    for file in file_list:
        # check the current list item "file_list" this is a file?
        if not os.path.isfile(file):
            print('"This is not file": {0}'.format(file))
            # это не файл, выполняем переход к следующей итерации цикла.
            continue
        if not file_is_needed(file):
            print('"This file is not needed": {0}'.format(file))
            # этот файл не нужен, выполняем переход к следующей итерации цикла.
            continue
        if skip_this_folder(file):
            print('"you need to skip this folder": {0}'.format(file))
            # этот файл не нужен, выполняем переход к следующей итерации цикла.
            continue
        print('обрабатываем файл: {0}'.format(file))
        curr_line = {}
        # file-name = file[2:]
        curr_line.update({'file-name': file[2:]})
        modtime = time.localtime(os.path.getmtime(file))
        curr_line['file_modify'] = time.strftime("%d-%m-%Y %H:%M:%S", modtime)
        curr_line.update({'Hash-MD5': get_hash_md5(file)})
        curr_line['author'] = author
        curr_line['validator'] = validator
        result_table.append(curr_line)
    return result_table


def save_table_to_txt_file(result_table, result_path):
    """ write 'result_table' to text file """

    f_name = result_path + '.txt'
    with open(f_name, "w", encoding='utf-8') as file:
        for line_dict in result_table:
            for key, val in line_dict.items():
                file.write('{}: {}\t'.format(key, val))
            file.write('\n')
        file.write(f'\n\t Исходный каталог = \"{source_path}\"\n')
        file.write(f'\n\t Количество файлов = \"{len(result_table)}\"\n')


def save_table_to_xlsx_file(result_table, result_path):
    """ write 'result_table' to xlsx file """

    wb = openpyxl.Workbook()
    wb.create_sheet(title='Список файлов', index=0)
    sheet = wb['Список файлов']
    # create a header row
    sheet['A1'] = '№№'
    sheet['B1'] = 'Имя файла относительно корня проекта'
    sheet.column_dimensions['B'].width = 100
    sheet['C1'] = 'дата последнего изменения'
    sheet['D1'] = 'Хэш-сумма (MD5) файла'
    # sheet.column_dimensions['C'].width = 30
    sheet['E1'] = 'Автор файла'
    # sheet.column_dimensions['D'].width = 20
    sheet['F1'] = 'Файл проверил'
    # sheet.column_dimensions['B'].width = 20

    row = 1
    for line_dict in result_table:
        row += 1
        col = 1
        cell = sheet.cell(row=row, column=col)
        cell.value = row-1

        for __, val in line_dict.items():
            col += 1
            cell = sheet.cell(row=row, column=col)
            cell.value = val

    col = 2
    cell = sheet.cell(row=row+2, column=col)
    cell.value = f'Исходный каталог = \"{source_path}\"'

    validator = auditors_list[1] if auditors_list else 'validator'
    cell = sheet.cell(row=row+4, column=col)
    cell.value = f'Сдал Руководитель проверки: {validator}'
    cell = sheet.cell(row=row+6, column=col)
    cell.value = f'Проверил Руководитель Контрольной Службы: {SKR_VALIDATOR}'

    f_name = result_path + '.xlsx'
    wb.save(f_name)
    

def process_create_file_table(source_path, result_path):
    """ Creates a list of all files as: file name, hash, author """

    file_list = get_file_list(source_path)
    if file_list is None:
        print('A list of all items in current directory is not created')
        return None
    result_table = create_table(file_list)
    save_table_to_txt_file(result_table, result_path)
    save_table_to_xlsx_file(result_table, result_path)

    # Удаляем Симвлолическую ссылку на диск V:
    remove_link_to_v()


if __name__ == '__main__':

    source_path = get_source_path()
    print(source_path)
    result_path = OUT_FILE_NAME
    process_create_file_table(
        source_path,
        result_path
    )
