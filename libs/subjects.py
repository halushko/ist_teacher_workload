def _add_value(string_number1, string_number2):
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

def update_subject(semestr, pdf, column, pdf_column):
    contract = pdf["Контракт"] == "Контракт"
    column_internal = column + ('.к' if contract else '.б')
    if pdf['Дисципліна'] not in semestr:
        add_subject(semestr, pdf, column, pdf_column)
    if column_internal not in semestr[pdf['Дисципліна']]:
        add_subject(semestr, pdf, column, pdf_column)
    else:
        semestr[pdf['Дисципліна']][column_internal] = _add_value(semestr[pdf['Дисципліна']][column_internal],
                                                                 pdf[pdf_column])


def add_subject(semestr, pdf, column, pdf_column):
    if pdf['Дисципліна'] not in semestr:
        semestr[pdf['Дисципліна']] = {
            'Факультет': 'ІОТ',
            'Шифр груп': pdf['Поток/Группа'],
            'Дисципліна': pdf['Дисципліна'],
            'Курс': pdf['Курс'],
        }

    semestr[pdf['Дисципліна']][column + '.б'] = ""
    semestr[pdf['Дисципліна']][column + '.к'] = ""
    update_subject(semestr, pdf, column, pdf_column)


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
        semestr[pdf['Дисципліна']][column_internal] = _add_value(semestr[pdf['Дисципліна']][column_internal],
                                                                 pdf[pdf_column])


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