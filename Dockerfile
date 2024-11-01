FROM python:3.10-slim

RUN apt-get update
RUN apt-get install -y poppler-utils cargo
RUN pip install python-telegram-bot openpyxl pdfplumber requests

WORKDIR /app
RUN mkdir ./files
RUN mkdir ./libs

COPY . .



CMD ["python", "main.py"]