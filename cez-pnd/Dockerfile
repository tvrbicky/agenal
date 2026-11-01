# Oficiální image Playwrightu – obsahuje Chromium i všechny systémové knihovny
FROM mcr.microsoft.com/playwright/python:v1.49.1-noble

WORKDIR /srv
# Carlito = metricky shodné s Calibri (písmo PDF reportu)
RUN apt-get update && apt-get install -y --no-install-recommends fonts-crosextra-carlito && rm -rf /var/lib/apt/lists/*
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY app ./app

ENV DATA_DIR=/data TZ=Europe/Prague PYTHONUNBUFFERED=1
VOLUME /data
EXPOSE 8080
CMD ["python", "-m", "app.web"]
