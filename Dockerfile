FROM python:3.11-slim

RUN apt-get update && apt-get install -y \
    libreoffice-calc \
    libreoffice-common \
    libreoffice-core \
    python3-uno \
    --no-install-recommends \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

ENV PORT=8080
CMD ["/bin/sh", "-c", "gunicorn --bind 0.0.0.0:${PORT} --timeout 300 app:app"]
