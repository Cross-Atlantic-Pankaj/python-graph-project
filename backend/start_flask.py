#!/usr/bin/env python3
"""
Simple Flask startup script for PM2
"""

import os
import sys

# Set environment variables
os.environ['PYTHONDONTWRITEBYTECODE'] = '1'
os.environ['FLASK_ENV'] = 'production'

# Import and run the Flask app
from app import create_app

if __name__ == '__main__':
    app = create_app()
    print("🚀 Starting Flask server...")
    app.run(debug=False, host='0.0.0.0', port=5001)
