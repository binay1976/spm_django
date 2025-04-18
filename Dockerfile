# Use the official Python 3.11 base image
FROM python:3.12.9-slim-bookworm

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Set work directory
WORKDIR /app

# Install system dependencies (optional, depending on your needs)
# RUN apt-get update && apt-get install -y \
#     build-essential \
#     && apt-get clean

# Install Python dependencies
COPY requirements.txt .
RUN pip install --upgrade pip
RUN pip install -r requirements.txt

# Copy project files
COPY . .

EXPOSE 8000

# Command to run your application (update this as per your main script)
CMD ["python", "manage.py", "runserver", "0.0.0:8000"]
