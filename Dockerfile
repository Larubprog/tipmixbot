# Use a lightweight Python image
FROM python:3.9-slim

# Set the working directory
WORKDIR /app

# Install system dependencies including Chrome and Chromedriver
RUN apt-get update && apt-get install -y \
    wget \
    curl \
    unzip \
    chromium \
    chromium-driver \
    libnss3 \
    libatk1.0-0 \
    libatk-bridge2.0-0 \
    libcups2 \
    libxkbcommon-x11-0 \
    libpango-1.0-0 \
    libxcomposite1 \
    libxrandr2 \
    libgbm1 \
    libasound2 \
    libxdamage1 \
    libxfixes3 \
    libx11-xcb1 \
    && rm -rf /var/lib/apt/lists/*

# Set environment variables for Selenium
ENV PATH="/usr/lib/chromium/:$PATH"

# Copy project files into the container
COPY . /app

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Install Playwright browsers
RUN playwright install --with-deps

# Expose a port if needed
EXPOSE 8000

# Set the default command to run the bot
CMD ["python", "-m", "src.main"]
