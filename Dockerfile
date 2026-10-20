# Resmi Python imajını kullan (Hafif ve hızlı)
FROM python:3.10-slim

# Sunucu içindeki çalışma klasörünü ayarla
WORKDIR /app

# Kütüphane listesini kopyala ve kur
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Projedeki tüm dosyaları kopyala
COPY . .

# Google Cloud Run varsayılan olarak 8080 portunu dinler
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8080"]