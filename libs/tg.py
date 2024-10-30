import os

import requests

from libs import file

echo_text = "Просто скиньте мені PDF з навантеженнями, вкажыть свої ПІБ і я згенерую Вам Excel"

start_text = "Привіт, я вмію парсити PDF з навантаженням! Просто скиньте мені PDF з навантеженнями, " + \
             "вкажіть свої ПІБ і я згенерую Вам Excel"

request_error_bad_file = "Не вдалося завантажити файл PDF. Можливо він пошкоджений"

send_file_text = "Результат обробки PDF.\n\nНе забудьте ввести інфорацію в сторінках 1, 10, виправити дати в сторінках " + \
                 "1 та 11 та внести номер протоколу (якщо є) на сторінці 11\n\nТакож перевірте сторінку 8-9, там може " + \
                 "бути відсутня інформація (зверніться до @halushko).\n\nМожуть бути присутні символи ### - зменшіть " + \
                 "розмір шрифту, або якось розшите сторінки. Зважте на те, що їх ще треба друкувати.\n\nЗ питаннями " + \
                 "щодо боту звертайтеся до @halushko. З питаннями щодо навантаження - до відповідальної за навантаження людини"


async def send_text(context, update, text):
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text=text
    )

async def send_document(context, update, file_path, text):
    await context.bot.send_document(
        chat_id=update.effective_chat.id,
        document=open(file_path, 'rb'),
        filename=os.path.basename(file_path),
        caption=text
)

async def download_file(update, context):
    message = update.message
    file_id = message.document.file_id
    file_object = await context.bot.get_file(file_id)
    file_name = message.document.file_name
    pdf_file_path = os.path.join(file.directory, file_name)

    url = file_object.file_path
    response = requests.get(url)
    if response.status_code == 200:
        with open(pdf_file_path, "wb") as f:
            f.write(response.content)
    else:
        await send_text(context, update, request_error_bad_file)
        return
    return file.rename_files_with_random_hex(pdf_file_path)