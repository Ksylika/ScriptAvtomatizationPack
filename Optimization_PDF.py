import os
import subprocess
import shutil
import openpyxl
import argparse
import datetime
import glob
import re


xml_content = """<?xml version="1.0" encoding="UTF-8"?>
<dublin_core schema="dc">
    <dcvalue element="identifier" qualifier="uri">https://nebdeti.ru/%s</dcvalue>
</dublin_core>
"""

contents = '%s	bundle:ORIGINAL'


def take_key(file):
    if os.path.exists(f'{os.getcwd()}\\{file}'):
        with open(args.key, 'r') as file:
            key = file.read()
            if '-sOutputFile=' in key:
                key = key.replace(re.search(r'-sOutputFile=(.*?)(?:\s|$)', key).group(1), '"%s"')
                return key
            else:
                return key + ' -sOutputFile="%s"'
    else:
        print(f'Файл {file} ненайден, будет использована дефолтный набор ключей для gswin64c')
        key = ('gswin64c -dPDFA -sProcessColorModel=DeviceCMYK -sDEVICE=pdfwrite -dPDFACompatibilityPolicy=1 '
               '-dPDFACompatibilityLevel=1.4 -dPDFSETTINGS=/ebook -dBATCH -dNOPAUSE -dQUIET -dFastWebView '
               '-sOutputFile="%s"')
        return key


def logg(*loggs):
    time_now = datetime.datetime.now()
    with open(f"{os.getcwd()}\\log.txt", "a") as file:
        file.write(f'{time_now.strftime("%Y-%m-%d\t%H:%M:%S")}\t{loggs[0]}\t{loggs[1]}\t{loggs[2]}\n')


def first_step(args):
    if not os.path.exists(f'{args.dout}\\NEW'):
        os.makedirs(f'{args.dout}\\NEW')
    if not os.path.exists(f'{args.dout}\\temp'):
        os.makedirs(f'{os.getcwd()}\\temp')

    folders = os.listdir(args.din)
    workbook = openpyxl.load_workbook(args.excel)
    sheet = workbook.active
    row_last = args.row
    for row in sheet.iter_rows(min_row=args.row, values_only=True):
        dirname, dspaceid, pdfuid = str(row[0]), row[1], row[2]
        if dirname and dspaceid and pdfuid is not None:
            logg(row_last, dirname, 'Начало обработки')
            if dirname in folders and len(glob.glob1(f'{args.din}\\{dirname}', "*.pdf")) == 1:
                folder_new = os.path.join(f'{args.dout}\\NEW', dirname)
                if not os.path.exists(folder_new):
                    os.makedirs(folder_new)
                pdf_file_name = [pdf_file for pdf_file in os.listdir(f'{args.din}\\{dirname}') if pdf_file.endswith('.pdf')]
                if not os.path.exists(f'{folder_new}\\{pdf_file_name[0]}'):
                    try:
                        shutil.copy(f'{args.din}\\{dirname}\\{pdf_file_name[0]}',
                                    os.path.join(f'{os.getcwd()}\\temp', f'temp-{pdf_file_name[0]}'))
                        subprocess.run(f'{args.key % pdf_file_name[0]} "temp-{pdf_file_name[0]}"', cwd=f'{os.getcwd()}\\temp',
                                   shell=True)

                        shutil.move(f'{os.getcwd()}\\temp\\{pdf_file_name[0]}',
                                    folder_new)
                        os.remove(f'{os.getcwd()}\\temp\\temp-{pdf_file_name[0]}')
                        with open(f"{folder_new}\\dublin_core.xml", "w", encoding="UTF8") as file:
                            file.write(xml_content % dspaceid)
                        with open(f"{folder_new}\\contents", "w", encoding="UTF8") as file:
                            file.write(contents % pdf_file_name[0])
                        with open(f"{folder_new}\\delete_contents", "w", encoding="UTF8") as file:
                            file.write(pdfuid)
                        logg(row_last, dirname, 'ОК')

                        row_last += 1
                    except:
                        logg(row_last, dirname, 'Команда GS введена не верно, проверьте правильность строки')
                        print('Команда GS введена не верно, проверьте правильность строки')
                        input('Нажмите ENTER, чтобы выйти.')
                        break
                else:
                    logg(row_last, dirname, 'Ошибка: pdf файл уже существует в NEW')

            else:
                if dirname not in folders:
                    logg(row_last, dirname, 'Ошибка: директория не найдена')
                elif len(glob.glob1(f'{args.din}\\{dirname}', "*.pdf")) == 0:
                    logg(row_last, dirname, 'Ошибка: файл пдф не найден')
                else:
                    logg(row_last, dirname, 'Ошибка: несколько файлов пдф')
                row_last += 1
        else:
            break


if __name__ == '__main__':
    parser = argparse.ArgumentParser()

    parser.add_argument("-din", type=str, default=f"{os.getcwd()}\\OLD",
                        help=f"Путь к каталогу данных. Дефолтное значение: {os.getcwd()}\\OLD")
    parser.add_argument("-dout", type=str, default=f'{os.getcwd()}',
                        help=f"Путь к новому катологу в котором будет создана папка NEW. Дефолтное значение {os.getcwd()}")
    parser.add_argument("-excel", type=str, default='df_compress_list.xlsx', help="Название файла excel,файл должен "
                                                                                  "находиться в одной директроии с "
                                                                                  "испольняемым скриптом. Дефолтное название df_compress_list.xlsx")
    parser.add_argument("-row", type=int, default=2,
                        help="Номер строки с которой начать считывать excel. Дефолтное значение 2")
    parser.add_argument("-key", type=str, default='key.txt', help="Название файла key, файл должен "
                                                                  "находиться в одной директроии с "
                                                                  "испольняемым скриптом. Дефолтное название key.txt")

    args = parser.parse_args()
    if not os.path.exists(args.din):
        print('Каталог с данными -din ненайден.')
    elif not os.path.exists(args.dout):
        print('Каталог для обработки -dout ненайден.')
    elif not os.path.exists(f'{os.getcwd()}\\{args.excel}'):
        print('Файл excel ненайден, необходимая запись example.xlsx.')
    else:
        args.key = take_key(args.key)
        first_step(args)
