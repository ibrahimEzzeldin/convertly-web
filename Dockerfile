FROM python:3.11-slim

# Install LibreOffice + fonts
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        libreoffice \
        libreoffice-writer \
        libreoffice-calc \
        fontconfig \
        fonts-liberation \
        fonts-dejavu \
        fonts-freefont-ttf \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

# Arabic + CJK fonts — installed after LibreOffice so a font-package
# name change doesn't block the whole build. Any package that can't be
# found is logged and skipped individually.
RUN apt-get update && \
    for pkg in fonts-noto fonts-noto-cjk fonts-noto-color-emoji \
               fonts-kacst fonts-hosny-amiri fonts-sil-scheherazade; do \
        apt-get install -y --no-install-recommends "$pkg" \
            || echo "WARN: skipped missing package $pkg"; \
    done && \
    apt-get clean && rm -rf /var/lib/apt/lists/* && \
    fc-cache -fv

WORKDIR /app

# Install Python deps (cached layer — only rebuilds when requirements.txt changes)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p uploads

EXPOSE 10000

CMD ["gunicorn", "app:app", "--bind", "0.0.0.0:10000", "--timeout", "120", "--workers", "2"]
