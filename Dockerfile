FROM python:3.10-slim
RUN apt-get update && apt-get install -y \
    build-essential \
    libffi-dev \
    && rm -rf /var/lib/apt/lists/*

RUN pip install --prefer-binary 'cryptography<41.0.0' python-telegram-bot openpyxl pdfplumber requests

WORKDIR /app
RUN mkdir ./files
RUN mkdir ./libs

COPY . .

CMD ["python", "main.py"]
