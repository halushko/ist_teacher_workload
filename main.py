import re
import pdfplumber
import openpyxl
import shutil
from telegram.ext import Application, CommandHandler, MessageHandler, filters
import os
import random
import requests

headers = ['№', 'Дисципліна', 'Курс', 'Ст/Зао/Ін', 'Поток/Группа', 'Кол.ст.', 'Лекції', 'Практ.зан.', 'Лаб.зан.',
           'Консульт.', 'М.кнтр.р.б', 'ДКР', 'РГР', 'Курс.пр/роб', 'Залік', 'Іспит', 'Керівництво Практикою',
           'Керівництво Маг.пр.', 'Керівництво Маг.Н.', 'Керівництво Аспірант.', 'Керівництво Бак.', 'Рецензув.',
           'Робота вДЕК',
           'Усього']
second_page = ['ДЕК', 'МагМП', 'МагМН', 'Аспірант', 'БАК']
directory = f'./app/files/'


def set_value(cell, value):
    cell.protection = openpyxl.styles.Protection(locked=False)
    try:
        result = float(value)
        if result != 0.0:
            cell.value = result
    except ValueError:
        cell.value = value


def add_value(string_number1, string_number2):
    string_number2 = string_number2.replace(",", ".")
    try:
        number1 = float(string_number1)
        try:
            number2 = float(string_number2)
            result = number1 + number2
            if result == 0.0:
                return ""
            else:
                return str(result)
        except ValueError:
            if string_number2 != "":
                return string_number1 + ", " + string_number2
            else:
                return string_number1
    except ValueError:
        if string_number1 != "":
            return string_number1 + ", " + string_number2
        else:
            return string_number2


def fill_xlsx_first_page(sheet, predmety_complex, start_row):
    i = start_row - 1
    for ii, (key, entry) in enumerate(predmety_complex.items(), start=start_row):
        i = i + 1
        if key in second_page:
            i = i - 1
            continue

        set_value(sheet[f"B{i}"], key)
        set_value(sheet[f"L{i}"], entry['Шифр груп'])

        set_value(sheet[f"M{i}"], entry['кількість студ.б'])
        set_value(sheet[f"N{i}"], entry['кількість студ.к'])

        set_value(sheet[f"Q{i}"], entry['Лекції.б'])
        set_value(sheet[f"S{i}"], entry['Лекції.к'])

        set_value(sheet[f"U{i}"], entry['Практичні заняття (семінари).б'])
        set_value(sheet[f"W{i}"], entry['Практичні заняття (семінари).к'])

        set_value(sheet[f"Y{i}"], entry['Лабораторні заняття.б'])
        set_value(sheet[f"AA{i}"], entry['Лабораторні заняття.к'])

        set_value(sheet[f"AC{i}"], entry['Екзамени.б'])
        set_value(sheet[f"AE{i}"], entry['Екзамени.к'])

        set_value(sheet[f"AK{i}"], entry['Заліки.б'])
        set_value(sheet[f"AM{i}"], entry['Заліки.к'])

        set_value(sheet[f"AO{i}"], entry['Контрольні роботи.б'])
        set_value(sheet[f"AQ{i}"], entry['Контрольні роботи.к'])

        set_value(sheet[f"AW{i}"], entry['Курсові проекти.б'])
        set_value(sheet[f"AY{i}"], entry['Курсові проекти.к'])

        set_value(sheet[f"BA{i}"], entry['РГР, РР, ГР.б'])
        set_value(sheet[f"BC{i}"], entry['РГР, РР, ГР.к'])

        set_value(sheet[f"BE{i}"], entry['ДКР.б'])
        set_value(sheet[f"BG{i}"], entry['ДКР.к'])

        set_value(sheet[f"BM{i}"], entry['Консультації.б'])
        set_value(sheet[f"BO{i}"], entry['Консультації.к'])


def fill_xlsx_second_page(sheet, predmety_complex, letters):
    for i, (key, entry) in enumerate(predmety_complex.items()):
        if key == 'ДЕК':
            set_value(sheet[f"{letters[0]}24"], entry['ДЕК.бб.кільк'])
            set_value(sheet[f"{letters[1]}24"], entry['ДЕК.бк.кільк'])
            set_value(sheet[f"{letters[2]}24"], entry['ДЕК.бб'])
            set_value(sheet[f"{letters[3]}24"], entry['ДЕК.бк'])
            if not entry['ДЕК.бб.кільк'] + entry['ДЕК.бк.кільк'] == "":
                set_value(sheet[f"{letters[4]}24"], entry['Факультет'])
                set_value(sheet[f"{letters[5]}24"], entry['Шифр груп'])
                set_value(sheet[f"{letters[6]}24"], entry['Курс'])
            set_value(sheet[f"{letters[0]}27"], entry['ДЕК.мпб.кільк'])
            set_value(sheet[f"{letters[1]}27"], entry['ДЕК.мпк.кільк'])
            set_value(sheet[f"{letters[2]}27"], entry['ДЕК.мпб'])
            set_value(sheet[f"{letters[3]}27"], entry['ДЕК.мпк'])
            if not entry['ДЕК.мпб.кільк'] + entry['ДЕК.мпк.кільк'] == "":
                set_value(sheet[f"{letters[4]}27"], entry['Факультет'])
                set_value(sheet[f"{letters[5]}27"], entry['Шифр груп'])
                set_value(sheet[f"{letters[6]}27"], entry['Курс'])
            set_value(sheet[f"{letters[0]}29"], entry['ДЕК.мнб.кільк'])
            set_value(sheet[f"{letters[1]}29"], entry['ДЕК.мнк.кільк'])
            set_value(sheet[f"{letters[2]}29"], entry['ДЕК.мнб'])
            set_value(sheet[f"{letters[3]}29"], entry['ДЕК.мнк'])
            if not entry['ДЕК.мнб.кільк'] + entry['ДЕК.мнк.кільк'] == "":
                set_value(sheet[f"{letters[4]}29"], entry['Факультет'])
                set_value(sheet[f"{letters[5]}29"], entry['Шифр груп'])
                set_value(sheet[f"{letters[6]}29"], entry['Курс'])
        elif key == 'БАК':
            set_value(sheet[f"{letters[0]}13"], entry['кількість студ.б'])
            set_value(sheet[f"{letters[1]}13"], entry['кількість студ.к'])
            set_value(sheet[f"{letters[2]}13"], entry['Керівництво.бб'])
            set_value(sheet[f"{letters[3]}13"], entry['Керівництво.бк'])
            if not entry['кількість студ.б'] + entry['кількість студ.к'] == "":
                set_value(sheet[f"{letters[4]}13"], entry['Факультет'])
                set_value(sheet[f"{letters[5]}13"], entry['Шифр груп'])
                set_value(sheet[f"{letters[6]}13"], entry['Курс'])
        elif key == 'МагМП':
            set_value(sheet[f"{letters[0]}14"], entry['кількість студ.б'])
            set_value(sheet[f"{letters[1]}14"], entry['кількість студ.к'])
            set_value(sheet[f"{letters[2]}14"], entry['Керівництво.мпб'])
            set_value(sheet[f"{letters[3]}14"], entry['Керівництво.мпк'])
            if not entry['кількість студ.б'] + entry['кількість студ.к'] == "":
                set_value(sheet[f"{letters[4]}14"], entry['Факультет'])
                set_value(sheet[f"{letters[5]}14"], entry['Шифр груп'])
                set_value(sheet[f"{letters[6]}14"], entry['Курс'])
        elif key == 'МагМН':
            set_value(sheet[f"{letters[0]}15"], entry['кількість студ.б'])
            set_value(sheet[f"{letters[1]}15"], entry['кількість студ.к'])
            set_value(sheet[f"{letters[2]}15"], entry['Керівництво.мнб'])
            set_value(sheet[f"{letters[3]}15"], entry['Керівництво.мнк'])
            if not entry['кількість студ.б'] + entry['кількість студ.к'] == "":
                set_value(sheet[f"{letters[4]}15"], entry['Факультет'])
                set_value(sheet[f"{letters[5]}15"], entry['Шифр груп'])
                set_value(sheet[f"{letters[6]}15"], entry['Курс'])
        elif key == 'Аспірант':
            set_value(sheet[f"{letters[0]}30"], entry['кількість студ.б'])
            set_value(sheet[f"{letters[1]}30"], entry['кількість студ.к'])
            set_value(sheet[f"{letters[2]}30"], entry['Керівництво.аб'])
            set_value(sheet[f"{letters[3]}30"], entry['Керівництво.ак'])
            if not entry['кількість студ.б'] + entry['кількість студ.к'] == "":
                set_value(sheet[f"{letters[4]}30"], entry['Факультет'])
                set_value(sheet[f"{letters[5]}30"], entry['Шифр груп'])
                set_value(sheet[f"{letters[6]}30"], entry['Курс'])


def process_table(table, semestr, contract):
    result = []
    previous_dict = {}
    for _, row in enumerate(table):
        my_dict = {}
        for index, item in enumerate(row):
            if index == 1 and item == '':
                my_dict[headers[index]] = previous_dict[headers[index]]
            else:
                my_dict[headers[index]] = item
        my_dict['Семестр'] = str(semestr)
        if contract == "к":
            my_dict["Контракт"] = "Контракт"
        else:
            my_dict["Контракт"] = "Бюджет"

        result.append(my_dict)
        previous_dict = my_dict
    return result


def rename_files_with_random_hex(file):
    name, extension = os.path.splitext(file)
    new_name = ''.join(
        random.choice('0123456789ABCDEF') + random.choice('0123456789ABCDEF') for _ in range(10)
    )
    new_filename = new_name + extension

    new_file_path = os.path.join(directory, new_filename)
    os.rename(file, new_file_path)
    return new_name, new_file_path


async def start_parsing(update, context):
    message = update.message
    file_id = message.document.file_id
    print("Шукаю для ПІБ " + message.caption + " (" + str(update.effective_chat.id) + ")")

    pib = ""
    match = re.search(r"^(.*?)(?: \(([\dкб,]+)\))?$", message.caption)
    numbers, letters, sem = [], [], []

    if match:
        pib = match.group(1).strip()  # Основной текст

        if match.group(2):
            pairs = re.findall(r"(\d)([кб])", match.group(2))
            sem = [int(s) for s, _ in pairs]
            letters = [letter for _, letter in pairs]
        else:
            sem = [1, 1, 2, 2]
            letters = ['б', 'к', 'б', 'к']

        print("Основной текст:", pib)
        print("Числа:", sem)
        print("Буквы:", letters)
    else:
        print("Строка не соответствует ожидаемому формату.")

    file = await context.bot.get_file(file_id)
    file_name = message.document.file_name
    pdf_file_path = os.path.join(directory, file_name)

    url = file.file_path
    response = requests.get(url)
    if response.status_code == 200:
        with open(pdf_file_path, "wb") as f:
            f.write(response.content)
    else:
        await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text="Не вдалося завантажити файл PDF, прикладіть до повідомлення файл з навантаженнями"
        )
        return

    new_name, pdf_path = rename_files_with_random_hex(pdf_file_path)
    predmety = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            page_text = page.extract_text()
            if pib in page_text:
                tables = page.extract_tables()
                if len(tables) - 2 != len(sem):
                    await context.bot.send_message(
                        chat_id=update.effective_chat.id,
                        text="У Вашому навантаженні відсутня частина інформації. Вкажіть які семестри відсутні, надіславши нове виправлене повідомлення з текстом:\n\n" + pib + " (1б,1к,2б,2к)\n\nВидаліть з цього тексту в скобках те навантажження, яке відсутнє. Цифра - номер семестру, літера - контракт, чи бюджет"
                    )
                    return
                if tables:
                    for i, _ in enumerate(sem):
                        res = process_table(tables[i + 1], sem[i], letters[i])
                        predmety.extend(res)
                break

    excel_path = 'Example.xlsx'
    copy_path = directory + pib.replace(" ", "") + new_name + '.xlsx'

    shutil.copy(excel_path, copy_path)
    wb = openpyxl.load_workbook(copy_path)
    sheet = wb["4-7"]
    if sheet.protection.sheet:
        sheet.protection.sheet = False

    predmety_complex = {"1": {}, "2": {}}

    for a in predmety:
        semestr = predmety_complex[a['Семестр']]
        dysciplina = semestr.get(a['Дисципліна'], {})
        is_contract = a["Контракт"] == "Контракт"

        if dysciplina:
            semestr[a['Дисципліна']]["Шифр груп"] = add_value(semestr[a['Дисципліна']]["Шифр груп"], a['Поток/Группа'])

            if is_contract:
                semestr[a['Дисципліна']]['кількість студ.к'] = add_value(semestr[a['Дисципліна']]['кількість студ.к'],
                                                                         a['Кол.ст.'])
                semestr[a['Дисципліна']]['Лекції.к'] = add_value(semestr[a['Дисципліна']]['Лекції.к'], a['Лекції'])
                semestr[a['Дисципліна']]['Практичні заняття (семінари).к'] = add_value(
                    semestr[a['Дисципліна']]['Практичні заняття (семінари).к'], a['Практ.зан.'])
                semestr[a['Дисципліна']]['Лабораторні заняття.к'] = add_value(
                    semestr[a['Дисципліна']]['Лабораторні заняття.к'], a['Лаб.зан.'])
                semestr[a['Дисципліна']]['Екзамени.к'] = add_value(semestr[a['Дисципліна']]['Екзамени.к'], a['Іспит'])
                semestr[a['Дисципліна']]['Заліки.к'] = add_value(semestr[a['Дисципліна']]['Заліки.к'], a['Залік'])
                semestr[a['Дисципліна']]['Контрольні роботи.к'] = add_value(
                    semestr[a['Дисципліна']]['Контрольні роботи.к'],
                    a['М.кнтр.р.б'])
                semestr[a['Дисципліна']]['Курсові проекти.к'] = add_value(semestr[a['Дисципліна']]['Курсові проекти.к'],
                                                                          a['Курс.пр/роб'])
                semestr[a['Дисципліна']]['РГР, РР, ГР.к'] = add_value(semestr[a['Дисципліна']]['РГР, РР, ГР.к'],
                                                                      a['РГР'])
                semestr[a['Дисципліна']]['ДКР.к'] = add_value(semestr[a['Дисципліна']]['ДКР.к'], a['ДКР'])
                semestr[a['Дисципліна']]['Консультації.к'] = add_value(semestr[a['Дисципліна']]['Консультації.к'],
                                                                       a['Консульт.'])

                semestr[a['Дисципліна']]['ДЕК.бк'] = add_value(semestr[a['Дисципліна']]['ДЕК.бк'],
                                                               a['Робота вДЕК'] if int(a['Курс']) != 6 else "")
                semestr[a['Дисципліна']]['ДЕК.мпк'] = add_value(semestr[a['Дисципліна']]['ДЕК.мпк'],
                                                                a['Робота вДЕК'] if int(a['Курс']) == 6 and 'МП' in a[
                                                                    'Поток/Группа'] else "")
                semestr[a['Дисципліна']]['ДЕК.мнк'] = add_value(semestr[a['Дисципліна']]['ДЕК.мнк'],
                                                                a['Робота вДЕК'] if int(a['Курс']) == 6 and 'МН' in a[
                                                                    'Поток/Группа'] else "")
                semestr[a['Дисципліна']]['ДЕК.бк.кільк'] = add_value(semestr[a['Дисципліна']]['ДЕК.бк.кільк'],
                                                                     a['Кол.ст.'] if int(a['Курс']) != 6 else "")
                semestr[a['Дисципліна']]['ДЕК.мпк.кільк'] = add_value(semestr[a['Дисципліна']]['ДЕК.мпк.кільк'],
                                                                      a['Кол.ст.'] if int(a['Курс']) == 6 and 'МП' in a[
                                                                          'Поток/Группа'] else "")
                semestr[a['Дисципліна']]['ДЕК.мнк.кільк'] = add_value(semestr[a['Дисципліна']]['ДЕК.мнк.кільк'],
                                                                      a['Кол.ст.'] if int(a['Курс']) == 6 and 'МН' in a[
                                                                          'Поток/Группа'] else "")

                semestr[a['Дисципліна']]['Керівництво.бк'] = add_value(semestr[a['Дисципліна']]['Керівництво.бк'],
                                                                       a['Керівництво Бак.'])
                semestr[a['Дисципліна']]['Керівництво.мнк'] = add_value(semestr[a['Дисципліна']]['Керівництво.мнк'],
                                                                        a['Керівництво Маг.Н.'])
                semestr[a['Дисципліна']]['Керівництво.мпк'] = add_value(semestr[a['Дисципліна']]['Керівництво.мпк'],
                                                                        a['Керівництво Маг.пр.'])
                semestr[a['Дисципліна']]['Керівництво.ак'] = add_value(semestr[a['Дисципліна']]['Керівництво.ак'],
                                                                       a['Керівництво Аспірант.'])
            else:
                semestr[a['Дисципліна']]['кількість студ.б'] = add_value(semestr[a['Дисципліна']]['кількість студ.б'],
                                                                         a['Кол.ст.'])
                semestr[a['Дисципліна']]['Лекції.б'] = add_value(semestr[a['Дисципліна']]['Лекції.б'], a['Лекції'])
                semestr[a['Дисципліна']]['Практичні заняття (семінари).б'] = add_value(
                    semestr[a['Дисципліна']]['Практичні заняття (семінари).б'], a['Практ.зан.'])
                semestr[a['Дисципліна']]['Лабораторні заняття.б'] = add_value(
                    semestr[a['Дисципліна']]['Лабораторні заняття.б'], a['Лаб.зан.'])
                semestr[a['Дисципліна']]['Екзамени.б'] = add_value(semestr[a['Дисципліна']]['Екзамени.б'], a['Іспит'])
                semestr[a['Дисципліна']]['Заліки.б'] = add_value(semestr[a['Дисципліна']]['Заліки.б'], a['Залік'])
                semestr[a['Дисципліна']]['Контрольні роботи.б'] = add_value(
                    semestr[a['Дисципліна']]['Контрольні роботи.б'],
                    a['М.кнтр.р.б'])
                semestr[a['Дисципліна']]['Курсові проекти.б'] = add_value(semestr[a['Дисципліна']]['Курсові проекти.б'],
                                                                          a['Курс.пр/роб'])
                semestr[a['Дисципліна']]['РГР, РР, ГР.б'] = add_value(semestr[a['Дисципліна']]['РГР, РР, ГР.б'],
                                                                      a['РГР'])
                semestr[a['Дисципліна']]['ДКР.б'] = add_value(semestr[a['Дисципліна']]['ДКР.б'], a['ДКР'])
                semestr[a['Дисципліна']]['Консультації.б'] = add_value(semestr[a['Дисципліна']]['Консультації.б'],
                                                                       a['Консульт.'])
                semestr[a['Дисципліна']]['ДЕК.бб'] = add_value(semestr[a['Дисципліна']]['ДЕК.бб'],
                                                               a['Робота вДЕК'] if int(a['Курс']) != 6 else "")
                semestr[a['Дисципліна']]['ДЕК.мпб'] = add_value(semestr[a['Дисципліна']]['ДЕК.мпб'],
                                                                a['Робота вДЕК'] if int(a['Курс']) == 6 and 'МП' in a[
                                                                    'Поток/Группа'] else "")
                semestr[a['Дисципліна']]['ДЕК.мнб'] = add_value(semestr[a['Дисципліна']]['ДЕК.мнб'],
                                                                a['Робота вДЕК'] if int(a['Курс']) == 6 and 'МН' in a[
                                                                    'Поток/Группа'] else "")
                semestr[a['Дисципліна']]['ДЕК.бб.кільк'] = add_value(semestr[a['Дисципліна']]['ДЕК.бб.кільк'],
                                                                     a['Кол.ст.'] if int(a['Курс']) != 6 else "")
                semestr[a['Дисципліна']]['ДЕК.мпб.кільк'] = add_value(semestr[a['Дисципліна']]['ДЕК.мпб.кільк'],
                                                                      a['Кол.ст.'] if int(a['Курс']) == 6 and 'МП' in a[
                                                                          'Поток/Группа'] else "")
                semestr[a['Дисципліна']]['ДЕК.мнб.кільк'] = add_value(semestr[a['Дисципліна']]['ДЕК.мнб.кільк'],
                                                                      a['Кол.ст.'] if int(a['Курс']) == 6 and 'МН' in a[
                                                                          'Поток/Группа'] else "")

                semestr[a['Дисципліна']]['Керівництво.бб'] = add_value(semestr[a['Дисципліна']]['Керівництво.бб'],
                                                                       a['Керівництво Бак.'])
                semestr[a['Дисципліна']]['Керівництво.мнб'] = add_value(semestr[a['Дисципліна']]['Керівництво.мнб'],
                                                                        a['Керівництво Маг.Н.'])
                semestr[a['Дисципліна']]['Керівництво.мпб'] = add_value(semestr[a['Дисципліна']]['Керівництво.мпб'],
                                                                        a['Керівництво Маг.пр.'])
                semestr[a['Дисципліна']]['Керівництво.аб'] = add_value(semestr[a['Дисципліна']]['Керівництво.аб'],
                                                                       a['Керівництво Аспірант.'])
        else:
            semestr[a['Дисципліна']] = {
                'Факультет': 'ІОТ',
                'Шифр груп': a['Поток/Группа'],
                'Дисципліна': a['Дисципліна'],
                'Курс': a['Курс'],

                'кількість студ.к': add_value("", a['Кол.ст.']) if is_contract else "",
                'Лекції.к': add_value("", a['Лекції']) if is_contract else "",
                'Практичні заняття (семінари).к': add_value("", a['Практ.зан.']) if is_contract else "",
                'Лабораторні заняття.к': add_value("", a['Лаб.зан.']) if is_contract else "",
                'Екзамени.к': add_value("", a['Іспит']) if is_contract else "",
                'Заліки.к': add_value("", a['Залік']) if is_contract else "",
                'Контрольні роботи.к': add_value("", a['М.кнтр.р.б']) if is_contract else "",
                'Курсові проекти.к': add_value("", a['Курс.пр/роб']) if is_contract else "",
                'РГР, РР, ГР.к': add_value("", a['РГР']) if is_contract else "",
                'ДКР.к': add_value("", a['ДКР']) if is_contract else "",
                'Консультації.к': add_value("", a['Консульт.']) if is_contract else "",

                'кількість студ.б': "" if is_contract else add_value("", a['Кол.ст.']),
                'Лекції.б': "" if is_contract else add_value("", a['Лекції']),
                'Практичні заняття (семінари).б': "" if is_contract else add_value("", a['Практ.зан.']),
                'Лабораторні заняття.б': "" if is_contract else add_value("", a['Лаб.зан.']),
                'Екзамени.б': "" if is_contract else add_value("", a['Іспит']),
                'Заліки.б': "" if is_contract else add_value("", a['Залік']),
                'Контрольні роботи.б': "" if is_contract else add_value("", a['М.кнтр.р.б']),
                'Курсові проекти.б': "" if is_contract else add_value("", a['Курс.пр/роб']),
                'РГР, РР, ГР.б': "" if is_contract else add_value("", a['РГР']),
                'ДКР.б': "" if is_contract else add_value("", a['ДКР']),
                'Консультації.б': "" if is_contract else add_value("", a['Консульт.']),

                'ДЕК.бк': add_value("", a['Робота вДЕК']) if is_contract and int(a['Курс']) != 6 else "",
                'ДЕК.мпк': add_value("", a['Робота вДЕК']) if is_contract and int(a['Курс']) == 6 and 'МП' in a[
                    'Поток/Группа'] else "",
                'ДЕК.мнк': add_value("", a['Робота вДЕК']) if is_contract and int(a['Курс']) == 6 and 'МН' in a[
                    'Поток/Группа'] else "",
                'ДЕК.бк.кільк': add_value("", a['Кол.ст.']) if is_contract and int(a['Курс']) != 6 else "",
                'ДЕК.мпк.кільк': add_value("", a['Кол.ст.']) if is_contract and int(a['Курс']) == 6 and 'МП' in a[
                    'Поток/Группа'] else "",
                'ДЕК.мнк.кільк': add_value("", a['Кол.ст.']) if is_contract and int(a['Курс']) == 6 and 'МН' in a[
                    'Поток/Группа'] else "",

                'ДЕК.бб': add_value("", a['Робота вДЕК']) if not is_contract and int(a['Курс']) != 6 else "",
                'ДЕК.мпб': add_value("", a['Робота вДЕК']) if not is_contract and int(a['Курс']) == 6 and 'МП' in a[
                    'Поток/Группа'] else "",
                'ДЕК.мнб': add_value("", a['Робота вДЕК']) if not is_contract and int(a['Курс']) == 6 and 'МН' in a[
                    'Поток/Группа'] else "",
                'ДЕК.бб.кільк': add_value("", a['Кол.ст.']) if not is_contract and int(a['Курс']) != 6 else "",
                'ДЕК.мпб.кільк': add_value("", a['Кол.ст.']) if not is_contract and int(a['Курс']) == 6 and 'МП' in a[
                    'Поток/Группа'] else "",
                'ДЕК.мнб.кільк': add_value("", a['Кол.ст.']) if not is_contract and int(a['Курс']) == 6 and 'МН' in a[
                    'Поток/Группа'] else "",

                'Керівництво.бк': add_value("", a['Керівництво Бак.']) if is_contract else "",
                'Керівництво.бб': add_value("", a['Керівництво Бак.']) if not is_contract else "",
                'Керівництво.мнк': add_value("", a['Керівництво Маг.Н.']) if is_contract else "",
                'Керівництво.мнб': add_value("", a['Керівництво Маг.Н.']) if not is_contract else "",
                'Керівництво.мпк': add_value("", a['Керівництво Маг.пр.']) if is_contract else "",
                'Керівництво.мпб': add_value("", a['Керівництво Маг.пр.']) if not is_contract else "",
                'Керівництво.ак': add_value("", a['Керівництво Аспірант.']) if is_contract else "",
                'Керівництво.аб': add_value("", a['Керівництво Аспірант.']) if not is_contract else "",
            }

    for i, (key, entry) in enumerate(predmety_complex.items()):
        for i, (key, entry) in enumerate(entry.items()):
            unique_items = list(dict.fromkeys(entry['Шифр груп'].split(", ")))
            entry['Шифр груп'] = ", ".join(unique_items)

    fill_xlsx_first_page(sheet, predmety_complex["1"], 10)
    fill_xlsx_first_page(sheet, predmety_complex["2"], 29)

    sheet = wb["8-9"]
    fill_xlsx_second_page(sheet, predmety_complex["1"], ['K', 'L', 'M', 'O', 'F', 'I', 'H'])
    fill_xlsx_second_page(sheet, predmety_complex["2"], ['U', 'V', 'W', 'Y', 'R', 'T', 'S'])

    wb.save(copy_path)
    print("Дані успішно записані в Excel.")

    await context.bot.send_document(
        chat_id=update.effective_chat.id,
        document=open(copy_path, 'rb'),
        filename=os.path.basename(copy_path),
        caption="Результат обробки PDF"
    )


async def start(update, context):
    await context.bot.send_message(chat_id=update.effective_chat.id,
                                   text="Привіт, я вмію парсити PDF з навантаженням! Просто скиньте мені PDF з навантеженнями, вкажыть свої ПІБ і я згенерую Вам Excel")


async def echo(update, context):
    await context.bot.send_message(chat_id=update.effective_chat.id,
                                   text="Просто скиньте мені PDF з навантеженнями, вкажыть свої ПІБ і я згенерую Вам Excel")


async def parse_pdf(update, context):
    await start_parsing(update, context)


def main() -> None:
    bot_token = '7554471229:AAGUNZOQpIdvCZ35DtfhfEW5rVu_OEHGdEw'  # os.getenv('BOT_TOKEN')
    application = Application.builder().token(bot_token).build()
    application.add_handler(CommandHandler("start", start))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, echo))
    application.add_handler(MessageHandler(filters.Document.PDF, parse_pdf))
    application.run_polling()


if __name__ == '__main__':
    main()
