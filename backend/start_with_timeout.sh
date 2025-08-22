#!/bin/bash

# Simple Gunicorn startup with increased timeout for batch processing
# Usage: ./start_with_timeout.sh

echo "🚀 Starting Graph Project API with 5-minute timeout..."

# Activate virtual environment if it exists
if [ -d "venv" ]; then
    source venv/bin/activate
fi

# Start Gunicorn with increased timeout
gunicorn \
    --bind 0.0.0.0:5001 \
    --workers 2 \
    --timeout 300 \
    --graceful-timeout 300 \
    --keep-alive 2 \
    --max-requests 1000 \
    --max-requests-jitter 50 \
    --preload \
    --log-level info \
    app:create_app()

echo "✅ API started with increased timeout limits!"
