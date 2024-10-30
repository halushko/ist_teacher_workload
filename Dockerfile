FROM python:3.10-slim

RUN pip install --upgrade pip
RUN apt-get update && apt-get install -y poppler-utils
RUN pip install python-telegram-bot openpyxl pdfplumber requests

WORKDIR /app
RUN mkdir ./files
RUN mkdir ./libs

COPY . .



CMD ["python", "main.py"]