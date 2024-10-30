FROM python:3.10-slim

RUN pip install --upgrade pip
RUN pip install python-telegram-bot openpyxl pdfplumber requests
RUN apt-get update && apt-get install -y poppler-utils

WORKDIR /app
RUN mkdir ./files
RUN mkdir ./libs

COPY . .



CMD ["python", "main.py"]