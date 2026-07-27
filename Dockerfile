# docx-service/Dockerfile
FROM python:3.11-slim

WORKDIR /app

# Instalar dependencias del sistema necesarias para python-docx y Pillow
RUN apt-get update && apt-get install -y \
    gcc \
    libxml2-dev \
    libxslt-dev \
    && rm -rf /var/lib/apt/lists/*

# Copiar e instalar dependencias Python
COPY requirements.txt .
RUN pip install --no-cache-dir --timeout=120 -r requirements.txt

# Copiar el código fuente
COPY . .

# Puerto del microservicio
EXPOSE 8001

# Comando de inicio: Gunicorn con workers uvicorn para manejar peticiones concurrentes
# -w 2 → 2 procesos independientes (cada uno puede atender una solicitud de DOCX al mismo tiempo)
# --timeout 120 → DOCX grandes pueden tardar; evita que Gunicorn mate el worker antes de terminar
CMD ["gunicorn", "app.main:app", \
     "-w", "2", \
     "-k", "uvicorn.workers.UvicornWorker", \
     "--bind", "0.0.0.0:8001", \
     "--timeout", "120", \
     "--access-logfile", "-"]