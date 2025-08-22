#!/usr/bin/env python3
"""
Startup script for the Flask application with Gunicorn
This script is designed to work with PM2
"""

import os
import sys
import subprocess

# Set environment variables
os.environ['PYTHONDONTWRITEBYTECODE'] = '1'
os.environ['FLASK_ENV'] = 'production'

def start_gunicorn():
    """Start Gunicorn with the Flask app"""
    try:
        # Gunicorn command with increased timeout for batch processing
        cmd = [
            'gunicorn',
            '--bind', '0.0.0.0:5001',
            '--timeout', '300',
            '--graceful-timeout', '300',
            '--workers', '2',
            '--preload',
            '--log-level', 'info',
            'app:create_app()'
        ]
        
        print("🚀 Starting Graph Project API with Gunicorn...")
        print(f"Command: {' '.join(cmd)}")
        
        # Start Gunicorn
        subprocess.run(cmd, check=True)
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Gunicorn failed to start: {e}")
        sys.exit(1)
    except KeyboardInterrupt:
        print("🛑 Shutting down gracefully...")
        sys.exit(0)

if __name__ == '__main__':
    start_gunicorn()
