#!/bin/bash

# Disable Python bytecode generation
export PYTHONDONTWRITEBYTECODE=1

# Install Python dependencies
pip install -r requirements.txt

# Create a simple WSGI entry point
echo "from app import create_app; app = create_app()" > vercel_app.py

echo "Build completed successfully!" 