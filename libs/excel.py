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
            _set_value(sheet[f"{sheet_columns[0]}24"], entry['ДЕК.кільк.б.б'])
            _set_value(sheet[f"{sheet_columns[1]}24"], entry['ДЕК.кільк.б.к'])
            _set_value(sheet[f"{sheet_columns[2]}24"], entry['ДЕК.б.б'])
            _set_value(sheet[f"{sheet_columns[3]}24"], entry['ДЕК.б.к'])
            if not entry['ДЕК.кільк.б.б'] + entry['ДЕК.кільк.б.к'] == "":
                _set_value(sheet[f"{sheet_columns[4]}24"], entry['Факультет'])
                _set_value(sheet[f"{sheet_columns[5]}24"], entry['Шифр груп'])
                _set_value(sheet[f"{sheet_columns[6]}24"], entry['Курс'])
            _set_value(sheet[f"{sheet_columns[0]}27"], entry['ДЕК.кільк.мп.б'])
            _set_value(sheet[f"{sheet_columns[1]}27"], entry['ДЕК.кільк.мп.к'])
            _set_value(sheet[f"{sheet_columns[2]}27"], entry['ДЕК.мп.б'])
            _set_value(sheet[f"{sheet_columns[3]}27"], entry['ДЕК.мп.к'])
            if not entry['ДЕК.кільк.мп.б'] + entry['ДЕК.кільк.мп.к'] == "":
                _set_value(sheet[f"{sheet_columns[4]}27"], entry['Факультет'])
                _set_value(sheet[f"{sheet_columns[5]}27"], entry['Шифр груп'])
                _set_value(sheet[f"{sheet_columns[6]}27"], entry['Курс'])
            _set_value(sheet[f"{sheet_columns[0]}29"], entry['ДЕК.кільк.мн.б'])
            _set_value(sheet[f"{sheet_columns[1]}29"], entry['ДЕК.кільк.мн.к'])
            _set_value(sheet[f"{sheet_columns[2]}29"], entry['ДЕК.мн.б'])
            _set_value(sheet[f"{sheet_columns[3]}29"], entry['ДЕК.мн.к'])
            if not entry['ДЕК.кільк.мн.б'] + entry['ДЕК.кільк.мн.к'] == "":
                _set_value(sheet[f"{sheet_columns[4]}29"], entry['Факультет'])
                _set_value(sheet[f"{sheet_columns[5]}29"], entry['Шифр груп'])
                _set_value(sheet[f"{sheet_columns[6]}29"], entry['Курс'])
        elif key == 'БАК':
            _set_value(sheet[f"{sheet_columns[0]}13"], entry['кількість студ.б'])
            _set_value(sheet[f"{sheet_columns[1]}13"], entry['кількість студ.к'])
            _set_value(sheet[f"{sheet_columns[2]}13"], entry['Керівництво.б.б'])
            _set_value(sheet[f"{sheet_columns[3]}13"], entry['Керівництво.б.к'])
            if not entry['кількість студ.б'] + entry['кількість студ.к'] == "":
                _set_value(sheet[f"{sheet_columns[4]}13"], entry['Факультет'])
                _set_value(sheet[f"{sheet_columns[5]}13"], entry['Шифр груп'])
                _set_value(sheet[f"{sheet_columns[6]}13"], entry['Курс'])
        elif key == 'МагМП':
            _set_value(sheet[f"{sheet_columns[0]}14"], entry['кількість студ.б'])
            _set_value(sheet[f"{sheet_columns[1]}14"], entry['кількість студ.к'])
            _set_value(sheet[f"{sheet_columns[2]}14"], entry['Керівництво.мп.б'])
            _set_value(sheet[f"{sheet_columns[3]}14"], entry['Керівництво.мп.к'])
            if not entry['кількість студ.б'] + entry['кількість студ.к'] == "":
                _set_value(sheet[f"{sheet_columns[4]}14"], entry['Факультет'])
                _set_value(sheet[f"{sheet_columns[5]}14"], entry['Шифр груп'])
                _set_value(sheet[f"{sheet_columns[6]}14"], entry['Курс'])
        elif key == 'МагМН':
            _set_value(sheet[f"{sheet_columns[0]}15"], entry['кількість студ.б'])
            _set_value(sheet[f"{sheet_columns[1]}15"], entry['кількість студ.к'])
            _set_value(sheet[f"{sheet_columns[2]}15"], entry['Керівництво.мн.б'])
            _set_value(sheet[f"{sheet_columns[3]}15"], entry['Керівництво.мн.к'])
            if not entry['кількість студ.б'] + entry['кількість студ.к'] == "":
                _set_value(sheet[f"{sheet_columns[4]}15"], entry['Факультет'])
                _set_value(sheet[f"{sheet_columns[5]}15"], entry['Шифр груп'])
                _set_value(sheet[f"{sheet_columns[6]}15"], entry['Курс'])
        elif key == 'Аспірант':
            _set_value(sheet[f"{sheet_columns[0]}30"], entry['кількість студ.б'])
            _set_value(sheet[f"{sheet_columns[1]}30"], entry['кількість студ.к'])
            _set_value(sheet[f"{sheet_columns[2]}30"], entry['Керівництво.а.б'])
            _set_value(sheet[f"{sheet_columns[3]}30"], entry['Керівництво.а.к'])
            if not entry['кількість студ.б'] + entry['кількість студ.к'] == "":
                _set_value(sheet[f"{sheet_columns[4]}30"], entry['Факультет'])
                _set_value(sheet[f"{sheet_columns[5]}30"], entry['Шифр груп'])
                _set_value(sheet[f"{sheet_columns[6]}30"], entry['Курс'])
        elif key == 'Рецензування':
            _set_value(sheet[f"{sheet_columns[0]}19"], entry['Рецензування.кільк.б.б'])
            _set_value(sheet[f"{sheet_columns[1]}19"], entry['Рецензування.кільк.б.к'])
            _set_value(sheet[f"{sheet_columns[2]}19"], entry['Рецензування.б.б'])
            _set_value(sheet[f"{sheet_columns[3]}19"], entry['Рецензування.б.к'])
            if not entry['Рецензування.кільк.б.б'] + entry['Рецензування.кільк.б.к'] == "":
                _set_value(sheet[f"{sheet_columns[4]}19"], entry['Факультет'])
                _set_value(sheet[f"{sheet_columns[5]}19"], entry['Шифр груп'])
                _set_value(sheet[f"{sheet_columns[6]}19"], entry['Курс'])
            _set_value(sheet[f"{sheet_columns[0]}20"], entry['Рецензування.кільк.мп.б'])
            _set_value(sheet[f"{sheet_columns[1]}20"], entry['Рецензування.кільк.мп.к'])
            _set_value(sheet[f"{sheet_columns[2]}20"], entry['Рецензування.мп.б'])
            _set_value(sheet[f"{sheet_columns[3]}20"], entry['Рецензування.мп.к'])
            if not entry['Рецензування.кільк.мп.б'] + entry['Рецензування.кільк.мп.к'] == "":
                _set_value(sheet[f"{sheet_columns[4]}20"], entry['Факультет'])
                _set_value(sheet[f"{sheet_columns[5]}20"], entry['Шифр груп'])
                _set_value(sheet[f"{sheet_columns[6]}20"], entry['Курс'])
            _set_value(sheet[f"{sheet_columns[0]}21"], entry['Рецензування.кільк.мн.б'])
            _set_value(sheet[f"{sheet_columns[1]}21"], entry['Рецензування.кільк.мн.к'])
            _set_value(sheet[f"{sheet_columns[2]}21"], entry['Рецензування.мн.б'])
            _set_value(sheet[f"{sheet_columns[3]}21"], entry['Рецензування.мн.к'])
            if not entry['Рецензування.кільк.мн.б'] + entry['Рецензування.кільк.мн.к'] == "":
                _set_value(sheet[f"{sheet_columns[4]}21"], entry['Факультет'])
                _set_value(sheet[f"{sheet_columns[5]}21"], entry['Шифр груп'])
                _set_value(sheet[f"{sheet_columns[6]}21"], entry['Курс'])