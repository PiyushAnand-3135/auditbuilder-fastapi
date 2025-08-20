# Step 1: Use Python base image
FROM python:3.11-slim

# Step 2: Set working directory
WORKDIR /app

# Step 3: Install system dependencies (gcc etc.)
RUN apt-get update && apt-get install -y gcc && rm -rf /var/lib/apt/lists/*

# Step 4: Copy requirements
COPY requirements.txt .

# Step 5: Create virtual environment + install dependencies
RUN python -m venv /opt/venv \
    && /opt/venv/bin/pip install --upgrade pip \
    && /opt/venv/bin/pip install -r requirements.txt

# Step 6: Ensure venv is used
ENV PATH="/opt/venv/bin:$PATH"

# Step 7: Copy app code
COPY . .

# Step 8: Expose FastAPI port
EXPOSE 8000

# Step 9: Start FastAPI (similar to your nixpacks config but no --reload in prod)
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
