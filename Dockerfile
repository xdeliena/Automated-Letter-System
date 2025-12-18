FROM python:3.10-slim

# Set environment variables to prevent interactive prompts and fix matplotlib
ENV DEBIAN_FRONTEND=noninteractive
ENV PYTHONUNBUFFERED=1
ENV MPLCONFIGDIR=/tmp/matplotlib
ENV XDG_CACHE_HOME=/tmp/cache

WORKDIR /app

# Create writable directories for matplotlib and other configs
RUN mkdir -p /tmp/matplotlib /tmp/cache && chmod 777 /tmp/matplotlib /tmp/cache

# Install only Python dependencies (no system packages needed for ReportLab)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Ensure output directory is writable
RUN mkdir -p /app/output && chmod 777 /app/output

EXPOSE 7860

CMD ["python", "app.py"]