FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Forçar modo produção — app falhará se APP_SECRET_KEY não estiver definida
ENV FLASK_ENV=production
ENV PORT=5000
EXPOSE 5000

CMD ["python", "app.py"]
