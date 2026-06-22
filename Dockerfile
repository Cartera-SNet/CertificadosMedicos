FROM python:3.11-slim

# Instalar LibreOffice y dependencias
RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-calc \
    fonts-dejavu \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Establecer directorio de trabajo
WORKDIR /app

# Copiar archivos de dependencias primero (para cachear)
COPY requirements.txt .

# Instalar dependencias Python
RUN pip install --no-cache-dir -r requirements.txt

# Copiar código de la aplicación
COPY . .

# Crear directorio de descargas
RUN mkdir -p /app/downloads

# Puerto que Railway asignará
ENV PORT=8080

# Comando de inicio con gunicorn
CMD ["sh", "-c", "gunicorn app:app --bind 0.0.0.0:${PORT:-8080} --workers 1 --timeout 600 --keep-alive 5"]