FROM python:3.9-slim

WORKDIR /app
RUN mkdir ./files
RUN mkdir ./libs

COPY . .

RUN pip install python-telegram-bot openpyxl pdfplumber requests

CMD ["python", "main.py"]