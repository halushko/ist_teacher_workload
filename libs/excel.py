import openpyxl

second_page = ['ДЕК', 'МагМП', 'МагМН', 'Аспірант', 'БАК', 'Рецензування']
template_path = './Example.xlsx'


def _set_value(cell, value):
    cell.protection = openpyxl.styles.Protection(locked=False)
    try:
        result = float(value)
        if result != 0.0:
            cell.value = result
    except ValueError:
        cell.value = value


def fill_first_page(excel_file_path, excel_subjects):
    wb = openpyxl.load_workbook(excel_file_path)

    sheet = wb["4-7"]
    _fill_xlsx_first_page(sheet, excel_subjects["1"], 10)
    _fill_xlsx_first_page(sheet, excel_subjects["2"], 29)

    wb.save(excel_file_path)


def fill_second_page(excel_file_path, excel_subjects):
    wb = openpyxl.load_workbook(excel_file_path)

    sheet = wb["8-9"]
    _fill_xlsx_second_page(sheet, excel_subjects["1"], ['K', 'L', 'M', 'O', 'F', 'I', 'H'])
    _fill_xlsx_second_page(sheet, excel_subjects["2"], ['U', 'V', 'W', 'Y', 'R', 'T', 'S'])

    wb.save(excel_file_path)


def _fill_xlsx_first_page(sheet, excel_subjects, start_row):
    if sheet.protection.sheet:
        sheet.protection.sheet = False
    i = start_row - 1
    for ii, (key, entry) in enumerate(excel_subjects.items(), start=start_row):
        i = i + 1
        if key in second_page:
            i = i - 1
            continue

        _set_value(sheet[f"B{i}"], key)
        _set_value(sheet[f"L{i}"], entry['Шифр груп'])

        _set_value(sheet[f"M{i}"], entry['кількість студ.б'])
        _set_value(sheet[f"N{i}"], entry['кількість студ.к'])

        _set_value(sheet[f"Q{i}"], entry['Лекції.б'])
        _set_value(sheet[f"S{i}"], entry['Лекції.к'])

        _set_value(sheet[f"U{i}"], entry['Практичні заняття (семінари).б'])
        _set_value(sheet[f"W{i}"], entry['Практичні заняття (семінари).к'])

        _set_value(sheet[f"Y{i}"], entry['Лабораторні заняття.б'])
        _set_value(sheet[f"AA{i}"], entry['Лабораторні заняття.к'])

        _set_value(sheet[f"AC{i}"], entry['Екзамени.б'])
        _set_value(sheet[f"AE{i}"], entry['Екзамени.к'])

        _set_value(sheet[f"AK{i}"], entry['Заліки.б'])
        _set_value(sheet[f"AM{i}"], entry['Заліки.к'])

        _set_value(sheet[f"AO{i}"], entry['Контрольні роботи.б'])
        _set_value(sheet[f"AQ{i}"], entry['Контрольні роботи.к'])

        _set_value(sheet[f"AW{i}"], entry['Курсові проекти.б'])
        _set_value(sheet[f"AY{i}"], entry['Курсові проекти.к'])

        _set_value(sheet[f"BA{i}"], entry['РГР, РР, ГР.б'])
        _set_value(sheet[f"BC{i}"], entry['РГР, РР, ГР.к'])

        _set_value(sheet[f"BE{i}"], entry['ДКР.б'])
        _set_value(sheet[f"BG{i}"], entry['ДКР.к'])

        _set_value(sheet[f"BM{i}"], entry['Консультації.б'])
        _set_value(sheet[f"BO{i}"], entry['Консультації.к'])


def _fill_xlsx_second_page(sheet, excel_subjects, sheet_columns):
    for i, (key, entry) in enumerate(excel_subjects.items()):
        if key == 'ДЕК':
            _fill_second_page_info(sheet, sheet_columns, '24', entry, 'ДЕК.кільк.б', 'ДЕК', 'б')
            _fill_second_page_info(sheet, sheet_columns, '27', entry, 'ДЕК.кільк.мп', 'ДЕК', 'мп')
            _fill_second_page_info(sheet, sheet_columns, '29', entry, 'ДЕК.кільк.мн', 'ДЕК', 'мн')
        elif key == 'БАК':
            _fill_second_page_info(sheet, sheet_columns, '13', entry, 'кількість студ', 'Керівництво', 'б')
        elif key == 'МагМП':
            _fill_second_page_info(sheet, sheet_columns, '14', entry, 'кількість студ', 'Керівництво', 'мп')
        elif key == 'МагМН':
            _fill_second_page_info(sheet, sheet_columns, '15', entry, 'кількість студ', 'Керівництво', 'мн')
        elif key == 'Аспірант':
            _fill_second_page_info(sheet, sheet_columns, '30', entry, 'кількість студ', 'Керівництво', 'а')
        elif key == 'Рецензування':
            _fill_second_page_info(sheet, sheet_columns, '19', entry, 'Рецензування.кільк.б', 'Рецензування', 'б')
            _fill_second_page_info(sheet, sheet_columns, '20', entry, 'Рецензування.кільк.мп', 'Рецензування', 'мп')
            _fill_second_page_info(sheet, sheet_columns, '21', entry, 'Рецензування.кільк.мн', 'Рецензування', 'мн')


def _fill_second_page_info(sheet, sheet_columns, sheet_row, entry, counter, source, who):
    _set_value(sheet[f"{sheet_columns[0]}{sheet_row}"], entry[counter + '.б'])
    _set_value(sheet[f"{sheet_columns[1]}{sheet_row}"], entry[counter + '.к'])
    _set_value(sheet[f"{sheet_columns[2]}{sheet_row}"], entry[source + '.' + who + '.б'])
    _set_value(sheet[f"{sheet_columns[3]}{sheet_row}"], entry[source + '.' + who + '.к'])
    if not entry[counter + '.б'] + entry[counter + '.к'] == "":
        _set_value(sheet[f"{sheet_columns[4]}{sheet_row}"], entry['Факультет'])
        _set_value(sheet[f"{sheet_columns[5]}{sheet_row}"], entry['Шифр груп'])
        _set_value(sheet[f"{sheet_columns[6]}{sheet_row}"], entry['Курс'])
