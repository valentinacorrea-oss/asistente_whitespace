FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Cloud Run inyecta PORT. Un solo worker: el candado de uso_diario vive en memoria del proceso.
CMD exec gunicorn --workers 1 --threads 8 --timeout 120 -b 0.0.0.0:${PORT:-8080} app:app
