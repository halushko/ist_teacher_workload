FROM python:3.10-slim

RUN apt-get update && apt-get install -y \
    build-essential \
    curl \
    && rm -rf /var/lib/apt/lists/* \
    && pip install --upgrade pip

RUN curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh -s -- -y && \
    export PATH="$HOME/.cargo/bin:$PATH" && \
    rustup update stable && \
    rustup default stable

ENV PATH="/root/.cargo/bin:${PATH}"

RUN pip install python-telegram-bot openpyxl pdfplumber requests

WORKDIR /app
RUN mkdir ./files
RUN mkdir ./libs

COPY . .

CMD ["python", "main.py"]
