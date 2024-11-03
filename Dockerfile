FROM python:3.10-alpine

RUN apk update && apk add --no-cache \
    build-base \
    libffi-dev \
    openssl-dev \
    cargo \
    gcc \
    musl-dev \
    python3-dev

RUN pip install --no-cache-dir --prefer-binary 'cryptography<41.0.0' python-telegram-bot openpyxl pdfplumber requests

WORKDIR /app
RUN mkdir ./files ./libs

COPY . .

CMD ["python", "main.py"]
