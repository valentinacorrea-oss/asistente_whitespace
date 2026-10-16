# Imagen base oficial de Python
FROM python:3.11-slim

# Evita que Python escriba archivos .pyc en el disco
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Directorio de trabajo dentro del contenedor
WORKDIR /app

# Copiar e instalar dependencias
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiar el resto del código del proyecto
COPY . .

# Puerto expuesto por defecto en Cloud Run
EXPOSE 8080

# Comando para ejecutar la aplicación (asumiendo que usas Flask/FastAPI con gunicorn)
CMD ["gunicorn", "--bind", "0.0.0.0:8080", "main:app"]