# Use official Python 3.11 slim image
FROM python:3.11-slim

# Set working directory inside container
WORKDIR /app

# Copy requirements first (for Docker layer caching)
COPY requirements.txt .

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt && \
    pip install --no-cache-dir waitress

# Copy all project files
COPY . .

# Expose port 8000
EXPOSE 8000

# Start the app using run_production.py
CMD ["python", "run_production.py"]
