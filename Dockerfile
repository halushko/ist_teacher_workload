# Используем Rust как базовый образ
FROM rust:latest

# Устанавливаем Python и другие необходимые инструменты
RUN apt-get update && apt-get install -y \
    python3 \
    python3-pip \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Установка Python-зависимостей
RUN pip3 install --prefer-binary python-telegram-bot openpyxl pdfplumber requests

# Создаем рабочую директорию и копируем файлы
WORKDIR /app
RUN mkdir ./files
RUN mkdir ./libs

COPY . .

# Запускаем приложение
CMD ["python3", "main.py"]
