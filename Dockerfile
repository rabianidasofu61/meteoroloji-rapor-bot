FROM python:3.11-slim

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONUNBUFFERED=1 \
    PLAYWRIGHT_BROWSERS_PATH=/ms-playwright \
    PORT=7860 \
    TZ=Europe/Istanbul

WORKDIR /app

# Türkçe/Arial font desteği (Liberation Sans / DejaVu) ve saat dilimi
RUN apt-get update && apt-get install -y --no-install-recommends \
    tzdata \
    fonts-liberation \
    fonts-dejavu-core \
    ca-certificates \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r /app/requirements.txt && \
    playwright install --with-deps chromium && \
    chmod -R 755 /ms-playwright

COPY . /app

# Hugging Face Spaces (uid 1000) ve Render için yazma izinleri
RUN mkdir -p "/app/Haftalık Rapor Bot Documents" "/app/Üretilen Raporlar" "/app/Haritalar" "/app/data" "/app/templates" && \
    chmod -R 777 /app

EXPOSE 7860

CMD ["python", "telegram_bot.py"]
