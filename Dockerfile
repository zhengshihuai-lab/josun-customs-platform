FROM python:3.11-slim

WORKDIR /app

# Install locales for Chinese support
RUN apt-get update && apt-get install -y --no-install-recommends \
    locales \
    && sed -i 's/# zh_TW.UTF-8/zh_TW.UTF-8/' /etc/locale.gen \
    && locale-gen \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

ENV LANG=zh_TW.UTF-8
ENV LC_ALL=zh_TW.UTF-8

# Install Python dependencies
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# Copy application files
COPY *.py ./

# Create data directory
RUN mkdir -p 平台數據/step_data 平台數據/uploads 平台數據/outputs

# Generate test data
RUN python3 建立範例資料.py || true

# Expose port (Render sets PORT env var automatically)
EXPOSE 10000

# Start the platform
CMD ["python3", "關務工作平台.py"]
