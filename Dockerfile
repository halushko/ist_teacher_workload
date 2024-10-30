FROM python:3.9-slim

RUN mkdir /app/files
RUN mkdir /app/libs

COPY . /app

RUN pip install python-telegram-bot openpyxl pdfplumber requests

CMD ["python", "main.py"]