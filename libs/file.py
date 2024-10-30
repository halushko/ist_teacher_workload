import os
import random
import shutil

from libs import excel

directory = f'./app/files/'

def rename_files_with_random_hex(file):
    name, extension = os.path.splitext(file)
    new_name = ''.join(
        random.choice('0123456789ABCDEF') + random.choice('0123456789ABCDEF') for _ in range(10)
    )
    new_filename = new_name + extension

    new_file_path = os.path.join(directory, new_filename)
    os.rename(file, new_file_path)
    return new_name, new_file_path

def create_response_xls_file(teacher_name, pdf_name):
    result_xls_path = directory + teacher_name.replace(" ", "") + pdf_name + '.xlsx'
    shutil.copy(excel.template_path, result_xls_path)
    return result_xls_path