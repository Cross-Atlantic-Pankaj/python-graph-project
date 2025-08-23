#!/usr/bin/env python3
"""
Startup script for the Flask application with Gunicorn
This script is designed to work with PM2
"""

import os
import sys
import subprocess
from pathlib import Path

# Set environment variables
os.environ['PYTHONDONTWRITEBYTECODE'] = '1'
os.environ['FLASK_ENV'] = 'production'

def find_gunicorn():
    """Find gunicorn executable"""
    # Check if we're in a virtual environment
    if hasattr(sys, 'real_prefix') or (hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix):
        # We're in a virtual environment
        venv_bin = Path(sys.prefix) / 'bin'
        gunicorn_path = venv_bin / 'gunicorn'
        if gunicorn_path.exists():
            return str(gunicorn_path)
    
    # Check common locations
    possible_paths = [
        '/home/ubuntu/Python-Graph-Project/backend/venv/bin/gunicorn',
        '/home/ubuntu/Python-Graph-Project/backend/.venv/bin/gunicorn',
        '/usr/local/bin/gunicorn',
        '/usr/bin/gunicorn'
    ]
    
    for path in possible_paths:
        if os.path.exists(path):
            return path
    
    return 'gunicorn'  # Fallback to PATH

def start_gunicorn():
    """Start Gunicorn with the Flask app"""
    try:
        gunicorn_path = find_gunicorn()
        print(f"🔍 Using Gunicorn at: {gunicorn_path}")
        
        # Gunicorn command with increased timeout for batch processing
        cmd = [
            gunicorn_path,
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
    except FileNotFoundError as e:
        print(f"❌ Gunicorn not found: {e}")
        print("💡 Try installing gunicorn: pip install gunicorn")
        sys.exit(1)

if __name__ == '__main__':
    start_gunicorn()
