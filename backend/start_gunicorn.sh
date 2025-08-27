#!/bin/bash

# Start script for Graph Project API with increased timeout limits
# This addresses the WORKER TIMEOUT issue in batch processing

echo "🚀 Starting Graph Project API with increased timeout limits..."

# Activate virtual environment if it exists
if [ -d "venv" ]; then
    echo "📦 Activating virtual environment..."
    source venv/bin/activate
fi

# Set environment variables
export PYTHONDONTWRITEBYTECODE=1
export FLASK_ENV=production

# Start Gunicorn with custom configuration
echo "🔧 Starting Gunicorn with 5-minute timeout for batch processing..."
gunicorn --config gunicorn.conf.py app:create_app()

echo "✅ Graph Project API started successfully!"


