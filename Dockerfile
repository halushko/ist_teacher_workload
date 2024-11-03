FROM alpine:3.18

RUN apk update && apk add --no-cache \
    python3 \
    py3-pip \
    build-base \
    libffi-dev \
    openssl-dev \
    cargo \
    gcc \
    musl-dev \
    python3-dev

RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir \
    'python-telegram-bot==20.3' \
    'openpyxl==3.0.10' \
    'pdfplumber==0.8.1' \
    'requests==2.28.1' \
    'cryptography>=36.0.0'

WORKDIR /app
COPY . .

CMD ["python3", "main.py"]
