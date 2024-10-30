import pdfplumber

from libs import subjects

pdf_headers = ['№', 'Дисципліна', 'Курс', 'Ст/Зао/Ін', 'Поток/Группа', 'Кол.ст.', 'Лекції', 'Практ.зан.', 'Лаб.зан.',
               'Консульт.', 'М.кнтр.р.б', 'ДКР', 'РГР', 'Курс.пр/роб', 'Залік', 'Іспит', 'Керівництво Практикою',
               'Керівництво Маг.пр.', 'Керівництво Маг.Н.', 'Керівництво Аспірант.', 'Керівництво Бак.', 'Рецензув.',
               'Робота вДЕК',
               'Усього']

request_error_text = "У Вашому навантаженні відсутня частина інформації. Вкажіть які семестри відсутні, " + \
                     "надіславши нове виправлене повідомлення з текстом:\n\n <ваш ПІБ> (1б,1к,2б,2к)\n\n" + \
                     "Видаліть з цього тексту в скобках те навантажження, яке відсутнє. Цифра - номер семестру," + \
                     " літера - контракт, чи бюджет"

def get_pdf_subjects(file_path, teacher, semester_numbers, semester_contracts):
    pdf_subjects = []
    with pdfplumber.open(file_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            page_text = page.extract_text()
            if teacher in page_text:
                tables = page.extract_tables()
                if len(tables) - 2 != len(semester_numbers):
                    return pdf_subjects, request_error_text
                if tables:
                    for i, _ in enumerate(semester_numbers):
                        res = _process_table(tables[i + 1], semester_numbers[i], semester_contracts[i])
                        pdf_subjects.extend(res)
                break
    return pdf_subjects, ""

def get_excel_subjects(pdf_subjects):
    excel_subjects = {"1": {}, "2": {}}

    for pdf in pdf_subjects:
        pdf_semester = excel_subjects[pdf['Семестр']]

        subjects.update_subject(pdf_semester, pdf, "Шифр груп", 'Поток/Группа')
        subjects.update_subject(pdf_semester, pdf, 'кількість студ', 'Кол.ст.')
        subjects.update_subject(pdf_semester, pdf, 'Лекції', 'Лекції')
        subjects.update_subject(pdf_semester, pdf, 'Практичні заняття (семінари)', 'Практ.зан.')
        subjects.update_subject(pdf_semester, pdf, 'Лабораторні заняття', 'Лаб.зан.')
        subjects.update_subject(pdf_semester, pdf, 'Екзамени', 'Іспит')
        subjects.update_subject(pdf_semester, pdf, 'Заліки', 'Залік')
        subjects.update_subject(pdf_semester, pdf, 'Контрольні роботи', 'М.кнтр.р.б')
        subjects.update_subject(pdf_semester, pdf, 'Курсові проекти', 'Курс.пр/роб')
        subjects.update_subject(pdf_semester, pdf, 'РГР, РР, ГР', 'РГР')
        subjects.update_subject(pdf_semester, pdf, 'ДКР', 'ДКР')
        subjects.update_subject(pdf_semester, pdf, 'Консультації', 'Консульт.')
        subjects.update_subject(pdf_semester, pdf, 'Консультації', 'Консульт.')

        subjects.update_dek(pdf_semester, pdf, 'ДЕК', 'Робота вДЕК')
        subjects.update_dek(pdf_semester, pdf, 'ДЕК.кільк', 'Кол.ст.')
        subjects.update_dek(pdf_semester, pdf, 'Рецензування.кільк', 'Рецензув.')
        subjects.update_dek(pdf_semester, pdf, 'Рецензування.кільк', 'Кол.ст.')

        subjects.update_subject(pdf_semester, pdf, 'Керівництво.б', 'Керівництво Бак.')
        subjects.update_subject(pdf_semester, pdf, 'Керівництво.мн', 'Керівництво Маг.Н.')
        subjects.update_subject(pdf_semester, pdf, 'Керівництво.мп', 'Керівництво Маг.пр.')
        subjects.update_subject(pdf_semester, pdf, 'Керівництво.а', 'Керівництво Аспірант.')

    for i, (key, entry) in enumerate(excel_subjects.items()):
        for i, (key, entry) in enumerate(entry.items()):
            unique_items = list(dict.fromkeys(entry['Шифр груп'].split(", ")))
            entry['Шифр груп'] = ", ".join(unique_items)
    return excel_subjects

def _process_table(table, semestr, contract):
    result = []
    previous_dict = {}
    for _, row in enumerate(table):
        my_dict = {}
        for index, item in enumerate(row):
            if index == 1 and item == '':
                my_dict[pdf_headers[index]] = previous_dict[pdf_headers[index]]
            else:
                my_dict[pdf_headers[index]] = item
        my_dict['Семестр'] = str(semestr)
        if contract == "к":
            my_dict["Контракт"] = "Контракт"
        else:
            my_dict["Контракт"] = "Бюджет"

        result.append(my_dict)
        previous_dict = my_dict
    return result