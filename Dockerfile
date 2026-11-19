# MarketMind AI - Cloud Deployment Container
FROM python:3.11-slim

WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements and install python packages
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application files
COPY . .

# Expose port (Cloud Run sets $PORT dynamically)
ENV PORT=8501
EXPOSE 8501

# Healthcheck
HEALTHCHECK CMD curl --fail http://localhost:${PORT}/_stcore/health || exit 1

# Launch Streamlit dashboard with dynamic port support
CMD ["sh", "-c", "streamlit run dashboard/app.py --server.port=${PORT} --server.address=0.0.0.0"]
