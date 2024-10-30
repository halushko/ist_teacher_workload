import os
import re

from telegram.ext import Application, CommandHandler, MessageHandler, filters

from libs import tg, pdf_parser, file, excel

async def start_parsing(update, context):
    message = update.message
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

    pdf_name, pdf_path = await tg.download_file(update, context)

    pdf_subjects, err = pdf_parser.get_pdf_subjects(pdf_path, pib, sem, letters)
    if not err == "":
        await tg.send_text(context, update, err)
        return

    excel_subjects = pdf_parser.get_excel_subjects(pdf_subjects)
    excel_file_path = file.create_response_xls_file(pib, pdf_name)

    excel.fill_first_page(excel_file_path, excel_subjects)
    excel.fill_second_page(excel_file_path, excel_subjects)

    print("Дані успішно записані в Excel.")

    await tg.send_document(context, update, excel_file_path, tg.send_file_text)

async def start(update, context):
    await tg.send_text(context, update, tg.start_text)

async def echo(update, context):
    await tg.send_text(context, update, tg.echo_text)

async def parse_pdf(update, context):
    await start_parsing(update, context)


def main() -> None:
    bot_token = os.getenv('BOT_TOKEN')
    application = Application.builder().token(bot_token).build()
    application.add_handler(CommandHandler("start", start))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, echo))
    application.add_handler(MessageHandler(filters.Document.PDF, parse_pdf))
    application.run_polling()


if __name__ == '__main__':
    main()
