#!/usr/bin/env python3
"""
Startup script for the Flask application that prevents __pycache__ generation
"""

import sys
import os

# Disable Python bytecode generation
sys.dont_write_bytecode = True
os.environ['PYTHONDONTWRITEBYTECODE'] = '1'

# Import and run the Flask app
from app import create_app

if __name__ == '__main__':
    app = create_app()
    print("🚀 Starting Flask server with __pycache__ disabled...")
    app.run(debug=True, host='0.0.0.0', port=5001)


