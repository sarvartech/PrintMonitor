FROM python:3.12-slim

# Set timezone and environment variables
ENV PYTHONUNBUFFERED=1 \
    TZ=Asia/Tashkent

WORKDIR /app

# Install system dependencies (tzdata for Tashkent timezone)
RUN apt-get update && apt-get install -y --no-install-recommends \
    tzdata \
    && rm -rf /var/lib/apt/lists/*

# Copy and install python requirements
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application files
COPY canon_server.py .
COPY kpi_export.py .
COPY kpi_template.xlsx .
COPY canon_dashboard.html .

# Expose HTTP dashboard port and Syslog UDP port
EXPOSE 5000/tcp
EXPOSE 5140/udp

# Run application
CMD ["python", "canon_server.py"]
