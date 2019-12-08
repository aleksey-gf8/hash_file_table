import os
import glob
import hashlib


def get_source_path():
    return os.getcwd()


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
    
    # Удаляем Символическую ссылку на диск V: если она существует
    remove_link_to_v()

    # Подключаем Симвлолическую ссылку диск V: указывает на "source_path"
    subst_command = 'subst V: "{0}"'.format(source_path)
    print(subst_command)
    ret_code =  os.system(subst_command)

    if ret_code != 0:
        print('Символ-ая ссылка НЕ СДЕЛАНА !! диск V: -> {0} Код возврата subst: {1}'.\
        format(source_path, ret_code))
        # завершаем и переходим к следующей итерации цикла по source_dir.
        return None
    else:
        print('Символ-ая ссылка СДЕЛАНА диск V: -> {0} Код возврата subst: {1}'.
        format(source_path, ret_code))

    # создаём рекурсивно список объектов внутри "V:\ (src_dir_level1)"
    return glob.glob('V:\\**\\*', recursive=True)


def create_table(file_list):
    """ create а table from a 'file_list' each row consists of:
    file name + file hash + author + validator
    """
    author = 'Boshirov R.'
    validator = 'Petrov A.'
    result_table =[]
    for file in file_list:
        # проверяем существует ли текущий файл из списка "file_list"
        if not os.path.isfile(file):
            print('"This is not file": {0}'.format(file))
            # это не файл, выполняем переход к следующей итерации цикла.
            continue
        print('обрабатываем файл: {0}'.format(file))

        curr_line = {}
        #file-name = file[2:]
        curr_line.update({'file-name':file[2:]})
        curr_line.update({'Hash-MD5':get_hash_md5(file)})
        curr_line['author'] = author
        curr_line['validator'] = validator
        result_table.append(curr_line)

    return result_table


def save_table_to_txt_file(result_table, result_path):
    """ write 'result_table' to text file """    
    
    with open(result_path, "w", encoding = 'utf-8') as file:
        for line_dict in result_table:
            for key, val in line_dict.items():
                file.write('{}: {}\t'.format(key, val))
            file.write('\n')
        file.write(f'\n\t Исходный каталог = \"{source_path}\"\n')
        file.write(f'\n\t Количество файлов = \"{len(result_table)}\"\n')


def process_create_file_table(source_path, result_path):
    """ Creates a list of all files as: file name, hash, author """

    file_list = get_file_list(source_path)
    result_table = create_table(file_list)
    save_table_to_txt_file(result_table, result_path)
    

if '__name__' == '__name__':

    source_path = get_source_path()
    result_path = r"result.txt" 

    process_create_file_table(
        source_path,
        result_path
    )

    # Удаляем Симвлолическую ссылку на диск V:
    remove_link_to_v()
    