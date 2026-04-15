FROM python:3.11-slim

# Install LibreOffice + fonts
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        libreoffice \
        libreoffice-writer \
        libreoffice-calc \
        fonts-liberation \
        fonts-dejavu \
        fonts-noto \
        fonts-noto-core \
        fonts-noto-cjk \
        fonts-noto-color-emoji \
        fonts-kacst \
        fonts-kacst-one \
        fonts-hosny-amiri \
        fonts-sil-scheherazade \
        fonts-arabeyes \
        fonts-farsiweb \
        fonts-freefont-ttf \
        fonts-croscore \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/* \
    && fc-cache -fv

WORKDIR /app

# Install Python deps (cached layer — only rebuilds when requirements.txt changes)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p uploads

EXPOSE 10000

CMD ["gunicorn", "app:app", "--bind", "0.0.0.0:10000", "--timeout", "120", "--workers", "2"]
