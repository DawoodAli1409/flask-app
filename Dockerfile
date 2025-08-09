FROM python:3.12-slim

# Install system dependencies
RUN apt-get update && apt-get install -y \
    libjpeg-dev \
    zlib1g-dev \
    && rm -rf /var/lib/apt/lists/*

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PORT=8080

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Health check endpoint
RUN echo "from flask import Flask; app = Flask(__name__); @app.route('/_ah/health') def health(): return 'Healthy'" > health.py

# Main execution command
CMD ["sh", "-c", "python main.py && gunicorn --bind :$PORT --workers 1 --threads 8 --timeout 0 health:app"]