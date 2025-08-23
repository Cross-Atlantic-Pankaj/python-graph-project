#!/bin/bash

# Start script that properly activates virtual environment
echo "🚀 Starting Graph Project API with virtual environment..."

# Navigate to the project directory
cd /home/ubuntu/Python-Graph-Project/backend

# Activate virtual environment
if [ -d "venv" ]; then
    echo "📦 Activating virtual environment: venv"
    source venv/bin/activate
elif [ -d ".venv" ]; then
    echo "📦 Activating virtual environment: .venv"
    source .venv/bin/activate
else
    echo "⚠️ No virtual environment found"
fi

# Set environment variables
export PYTHONDONTWRITEBYTECODE=1
export FLASK_ENV=production

# Check if gunicorn is available
if command -v gunicorn &> /dev/null; then
    echo "✅ Gunicorn found at: $(which gunicorn)"
else
    echo "❌ Gunicorn not found. Installing..."
    pip install gunicorn
fi

# Start Gunicorn with increased timeout
echo "🔧 Starting Gunicorn with 5-minute timeout for batch processing..."
gunicorn \
    --bind 0.0.0.0:5001 \
    --timeout 300 \
    --graceful-timeout 300 \
    --workers 2 \
    --preload \
    --log-level info \
    app:create_app()

echo "✅ API started with increased timeout limits!"
