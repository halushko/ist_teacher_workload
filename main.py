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
second_page = ['ДЕК', 'МагМП', 'МагМН', 'Аспірант', 'БАК', 'Рецензування']
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
            set_value(sheet[f"{letters[0]}24"], entry['ДЕК.кільк.б.б'])
            set_value(sheet[f"{letters[1]}24"], entry['ДЕК.кільк.б.к'])
            set_value(sheet[f"{letters[2]}24"], entry['ДЕК.б.б'])
            set_value(sheet[f"{letters[3]}24"], entry['ДЕК.б.к'])
            if not entry['ДЕК.кільк.б.б'] + entry['ДЕК.кільк.б.к'] == "":
                set_value(sheet[f"{letters[4]}24"], entry['Факультет'])
                set_value(sheet[f"{letters[5]}24"], entry['Шифр груп'])
                set_value(sheet[f"{letters[6]}24"], entry['Курс'])
            set_value(sheet[f"{letters[0]}27"], entry['ДЕК.кільк.мп.б'])
            set_value(sheet[f"{letters[1]}27"], entry['ДЕК.кільк.мп.к'])
            set_value(sheet[f"{letters[2]}27"], entry['ДЕК.мп.б'])
            set_value(sheet[f"{letters[3]}27"], entry['ДЕК.мп.к'])
            if not entry['ДЕК.кільк.мп.б'] + entry['ДЕК.кільк.мп.к'] == "":
                set_value(sheet[f"{letters[4]}27"], entry['Факультет'])
                set_value(sheet[f"{letters[5]}27"], entry['Шифр груп'])
                set_value(sheet[f"{letters[6]}27"], entry['Курс'])
            set_value(sheet[f"{letters[0]}29"], entry['ДЕК.кільк.мн.б'])
            set_value(sheet[f"{letters[1]}29"], entry['ДЕК.кільк.мн.к'])
            set_value(sheet[f"{letters[2]}29"], entry['ДЕК.мн.б'])
            set_value(sheet[f"{letters[3]}29"], entry['ДЕК.мн.к'])
            if not entry['ДЕК.кільк.мн.б'] + entry['ДЕК.кільк.мн.к'] == "":
                set_value(sheet[f"{letters[4]}29"], entry['Факультет'])
                set_value(sheet[f"{letters[5]}29"], entry['Шифр груп'])
                set_value(sheet[f"{letters[6]}29"], entry['Курс'])
        elif key == 'БАК':
            set_value(sheet[f"{letters[0]}13"], entry['кількість студ.б'])
            set_value(sheet[f"{letters[1]}13"], entry['кількість студ.к'])
            set_value(sheet[f"{letters[2]}13"], entry['Керівництво.б.б'])
            set_value(sheet[f"{letters[3]}13"], entry['Керівництво.б.к'])
            if not entry['кількість студ.б'] + entry['кількість студ.к'] == "":
                set_value(sheet[f"{letters[4]}13"], entry['Факультет'])
                set_value(sheet[f"{letters[5]}13"], entry['Шифр груп'])
                set_value(sheet[f"{letters[6]}13"], entry['Курс'])
        elif key == 'МагМП':
            set_value(sheet[f"{letters[0]}14"], entry['кількість студ.б'])
            set_value(sheet[f"{letters[1]}14"], entry['кількість студ.к'])
            set_value(sheet[f"{letters[2]}14"], entry['Керівництво.мп.б'])
            set_value(sheet[f"{letters[3]}14"], entry['Керівництво.мп.к'])
            if not entry['кількість студ.б'] + entry['кількість студ.к'] == "":
                set_value(sheet[f"{letters[4]}14"], entry['Факультет'])
                set_value(sheet[f"{letters[5]}14"], entry['Шифр груп'])
                set_value(sheet[f"{letters[6]}14"], entry['Курс'])
        elif key == 'МагМН':
            set_value(sheet[f"{letters[0]}15"], entry['кількість студ.б'])
            set_value(sheet[f"{letters[1]}15"], entry['кількість студ.к'])
            set_value(sheet[f"{letters[2]}15"], entry['Керівництво.мн.б'])
            set_value(sheet[f"{letters[3]}15"], entry['Керівництво.мн.к'])
            if not entry['кількість студ.б'] + entry['кількість студ.к'] == "":
                set_value(sheet[f"{letters[4]}15"], entry['Факультет'])
                set_value(sheet[f"{letters[5]}15"], entry['Шифр груп'])
                set_value(sheet[f"{letters[6]}15"], entry['Курс'])
        elif key == 'Аспірант':
            set_value(sheet[f"{letters[0]}30"], entry['кількість студ.б'])
            set_value(sheet[f"{letters[1]}30"], entry['кількість студ.к'])
            set_value(sheet[f"{letters[2]}30"], entry['Керівництво.а.б'])
            set_value(sheet[f"{letters[3]}30"], entry['Керівництво.а.к'])
            if not entry['кількість студ.б'] + entry['кількість студ.к'] == "":
                set_value(sheet[f"{letters[4]}30"], entry['Факультет'])
                set_value(sheet[f"{letters[5]}30"], entry['Шифр груп'])
                set_value(sheet[f"{letters[6]}30"], entry['Курс'])
        elif key == 'Рецензування':
            set_value(sheet[f"{letters[0]}19"], entry['Рецензування.кільк.б.б'])
            set_value(sheet[f"{letters[1]}19"], entry['Рецензування.кільк.б.к'])
            set_value(sheet[f"{letters[2]}19"], entry['Рецензування.б.б'])
            set_value(sheet[f"{letters[3]}19"], entry['Рецензування.б.к'])
            if not entry['Рецензування.кільк.б.б'] + entry['Рецензування.кільк.б.к'] == "":
                set_value(sheet[f"{letters[4]}19"], entry['Факультет'])
                set_value(sheet[f"{letters[5]}19"], entry['Шифр груп'])
                set_value(sheet[f"{letters[6]}19"], entry['Курс'])
            set_value(sheet[f"{letters[0]}20"], entry['Рецензування.кільк.мп.б'])
            set_value(sheet[f"{letters[1]}20"], entry['Рецензування.кільк.мп.к'])
            set_value(sheet[f"{letters[2]}20"], entry['Рецензування.мп.б'])
            set_value(sheet[f"{letters[3]}20"], entry['Рецензування.мп.к'])
            if not entry['Рецензування.кільк.мп.б'] + entry['Рецензування.кільк.мп.к'] == "":
                set_value(sheet[f"{letters[4]}20"], entry['Факультет'])
                set_value(sheet[f"{letters[5]}20"], entry['Шифр груп'])
                set_value(sheet[f"{letters[6]}20"], entry['Курс'])
            set_value(sheet[f"{letters[0]}21"], entry['Рецензування.кільк.мн.б'])
            set_value(sheet[f"{letters[1]}21"], entry['Рецензування.кільк.мн.к'])
            set_value(sheet[f"{letters[2]}21"], entry['Рецензування.мн.б'])
            set_value(sheet[f"{letters[3]}21"], entry['Рецензування.мн.к'])
            if not entry['Рецензування.кільк.мн.б'] + entry['Рецензування.кільк.мн.к'] == "":
                set_value(sheet[f"{letters[4]}21"], entry['Факультет'])
                set_value(sheet[f"{letters[5]}21"], entry['Шифр груп'])
                set_value(sheet[f"{letters[6]}21"], entry['Курс'])

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


def update_disciplina(semestr, pdf, column, pdf_column):
    contract = pdf["Контракт"] == "Контракт"
    column_internal = column + ('.к' if contract else '.б')
    if pdf['Дисципліна'] not in semestr:
        add_disciplina(semestr, pdf, column, pdf_column)
    if column_internal not in semestr[pdf['Дисципліна']]:
        add_disciplina(semestr, pdf, column, pdf_column)
    else:
        semestr[pdf['Дисципліна']][column_internal] = add_value(semestr[pdf['Дисципліна']][column_internal], pdf[pdf_column])


def add_disciplina(semestr, pdf, column, pdf_column):
    if pdf['Дисципліна'] not in semestr:
        semestr[pdf['Дисципліна']] = {
            'Факультет': 'ІОТ',
            'Шифр груп': pdf['Поток/Группа'],
            'Дисципліна': pdf['Дисципліна'],
            'Курс': pdf['Курс'],
        }

    semestr[pdf['Дисципліна']][column + '.б'] = ""
    semestr[pdf['Дисципліна']][column + '.к'] = ""
    update_disciplina(semestr, pdf, column, pdf_column)

def update_dek(semestr, pdf, column, pdf_column):
    contract = pdf["Контракт"] == "Контракт"
    suffix = '.'
    if int(pdf['Курс']) != 6:
        suffix = suffix + 'б'
    elif 'МП' in pdf['Поток/Группа']:
        suffix = suffix + 'мп'
    elif 'МН' in pdf['Поток/Группа']:
        suffix = suffix + 'мн'
    else:
        return
    suffix = suffix + ('.к' if contract else '.б')
    column_internal = column + suffix

    if pdf['Дисципліна'] not in semestr:
        add_dek(semestr, pdf, column, pdf_column)
    if column_internal not in semestr[pdf['Дисципліна']]:
        add_dek(semestr, pdf, column, pdf_column)
    else:
        semestr[pdf['Дисципліна']][column_internal] = add_value(semestr[pdf['Дисципліна']][column_internal], pdf[pdf_column])

def add_dek(semestr, pdf, column, pdf_column):
    if pdf['Дисципліна'] not in semestr:
        semestr[pdf['Дисципліна']] = {
            'Факультет': 'ІОТ',
            'Шифр груп': pdf['Поток/Группа'],
            'Дисципліна': pdf['Дисципліна'],
            'Курс': pdf['Курс'],
        }
    semestr[pdf['Дисципліна']][column + '.б.б'] = ""
    semestr[pdf['Дисципліна']][column + '.б.к'] = ""
    semestr[pdf['Дисципліна']][column + '.мп.б'] = ""
    semestr[pdf['Дисципліна']][column + '.мп.к'] = ""
    semestr[pdf['Дисципліна']][column + '.мн.б'] = ""
    semestr[pdf['Дисципліна']][column + '.мн.к'] = ""
    
    update_dek(semestr, pdf, column, pdf_column)

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

    excel_info = {"1": {}, "2": {}}

    for pdf in predmety:
        semestr = excel_info[pdf['Семестр']]

        update_disciplina(semestr, pdf, "Шифр груп", 'Поток/Группа')
        update_disciplina(semestr, pdf, 'кількість студ', 'Кол.ст.')
        update_disciplina(semestr, pdf, 'Лекції', 'Лекції')
        update_disciplina(semestr, pdf, 'Практичні заняття (семінари)', 'Практ.зан.')
        update_disciplina(semestr, pdf, 'Лабораторні заняття', 'Лаб.зан.')
        update_disciplina(semestr, pdf, 'Екзамени', 'Іспит')
        update_disciplina(semestr, pdf, 'Заліки', 'Залік')
        update_disciplina(semestr, pdf, 'Контрольні роботи', 'М.кнтр.р.б')
        update_disciplina(semestr, pdf, 'Курсові проекти', 'Курс.пр/роб')
        update_disciplina(semestr, pdf, 'РГР, РР, ГР', 'РГР')
        update_disciplina(semestr, pdf, 'ДКР', 'ДКР')
        update_disciplina(semestr, pdf, 'Консультації', 'Консульт.')
        update_disciplina(semestr, pdf, 'Консультації', 'Консульт.')

        update_dek(semestr, pdf, 'ДЕК', 'Робота вДЕК')
        update_dek(semestr, pdf, 'ДЕК.кільк', 'Кол.ст.')
        update_dek(semestr, pdf, 'Рецензування.кільк', 'Рецензув.')
        update_dek(semestr, pdf, 'Рецензування.кільк', 'Кол.ст.')

        update_disciplina(semestr, pdf, 'Керівництво.б', 'Керівництво Бак.')
        update_disciplina(semestr, pdf, 'Керівництво.мн', 'Керівництво Маг.Н.')
        update_disciplina(semestr, pdf, 'Керівництво.мп', 'Керівництво Маг.пр.')
        update_disciplina(semestr, pdf, 'Керівництво.а', 'Керівництво Аспірант.')

    for i, (key, entry) in enumerate(excel_info.items()):
        for i, (key, entry) in enumerate(entry.items()):
            unique_items = list(dict.fromkeys(entry['Шифр груп'].split(", ")))
            entry['Шифр груп'] = ", ".join(unique_items)

    fill_xlsx_first_page(sheet, excel_info["1"], 10)
    fill_xlsx_first_page(sheet, excel_info["2"], 29)

    sheet = wb["8-9"]
    fill_xlsx_second_page(sheet, excel_info["1"], ['K', 'L', 'M', 'O', 'F', 'I', 'H'])
    fill_xlsx_second_page(sheet, excel_info["2"], ['U', 'V', 'W', 'Y', 'R', 'T', 'S'])

    wb.save(copy_path)
    print("Дані успішно записані в Excel.")

    await context.bot.send_document(
        chat_id=update.effective_chat.id,
        document=open(copy_path, 'rb'),
        filename=os.path.basename(copy_path),
        caption="Результат обробки PDF.\n\nНе забудьте ввести інфорацію в сторінках 1, 10, виправити дати в сторінках 1 та 11 та внести номер протоколу (якщо є) на сторінці 11\n\nТакож перевірте сторінку 8-9, там може бути відсутня інформація (зверніться до @halushko).\n\nМожуть бути присутні символи ### - зменшіть розмір шрифту, або якось розшите сторінки. Зважте на те, що їх ще треба друкувати.\n\nЗ питаннями щодо боту звертайтеся до @halushko. З питаннями щодо навантаження - до відповідальної за навантаження людини"
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
