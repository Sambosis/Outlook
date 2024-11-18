# Use Python 3.9 slim image
FROM python:3.9-slim

# Set working directory
WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements first to leverage Docker cache
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the application
COPY . .

# Create the gpg2 directory
RUN mkdir -p /app/gpg2

# Set environment variables
ENV FLASK_APP=app.py
ENV FLASK_ENV=development
ENV EMAIL_DIR=/app/gpg2

# Expose port 5000
EXPOSE 5000

# Run the application
CMD ["python", "app.py"]