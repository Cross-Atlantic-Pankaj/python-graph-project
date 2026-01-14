import sys
import os
import logging

# Disable Python bytecode generation to prevent __pycache__ files
sys.dont_write_bytecode = True

# Configure logging to reduce verbose output
logging.basicConfig(
    level=logging.WARNING,  # Change from INFO to WARNING to reduce verbosity
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

# Suppress matplotlib font warnings
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='matplotlib')

# Suppress matplotlib font manager warnings
try:
    import matplotlib
    matplotlib.set_loglevel('error')  # Only show errors, not warnings
except ImportError:
    pass

from flask import Flask, send_from_directory, jsonify, current_app
import re
try:
    import matplotlib.pyplot as plt
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False
    print("Warning: matplotlib not available - chart functionality will be limited")
from openpyxl import load_workbook
from docx import Document
from docx.shared import Inches
from flask import current_app
from flask_pymongo import PyMongo
from flask_login import LoginManager
from flask_cors import CORS
import os
from bson.objectid import ObjectId
from dotenv import load_dotenv
import socket

load_dotenv()

mongo = PyMongo() # Define the PyMongo instance globally
login_manager = LoginManager()

def create_app():
    app = Flask(__name__, static_folder='../frontend-react/build', static_url_path='')
    
    # Load configuration based on environment
    config_name = os.environ.get('FLASK_ENV', 'production')
    from config import config
    app.config.from_object(config[config_name])
    
    # Configure app logging
    if not app.debug:
        app.logger.setLevel(logging.WARNING)
    
    # Load MONGO_URI from environment variable
    app.config["MONGO_URI"] = os.environ["MONGO_URI"]
    # Configure cookies to work over HTTP on AWS. When you move to HTTPS, set
    # SESSION_COOKIE_SECURE=True and SESSION_COOKIE_SAMESITE='None'.
    app.config['SESSION_COOKIE_SECURE'] = False
    # For cross-origin on HTTP: Try 'None' as string (some browsers may accept it even without Secure)
    # If this doesn't work, the browser is blocking it and you'll need HTTPS
    # Alternative: Set to None (Python None) to omit SameSite attribute entirely
    app.config['SESSION_COOKIE_SAMESITE'] = 'None'  # String 'None' - may work on some browsers even without Secure
    app.config['SESSION_COOKIE_HTTPONLY'] = True
    app.config['SESSION_COOKIE_DOMAIN'] = None
    # Get allowed origins from environment variable or use defaults
    # For deployment, set CORS_ORIGINS env var with comma-separated origins
    # Example: CORS_ORIGINS=http://localhost:3000,http://localhost:3001
    cors_origins_env = os.environ.get('CORS_ORIGINS', '')
    if cors_origins_env:
        # Split by comma and strip whitespace
        allowed_origins = [origin.strip() for origin in cors_origins_env.split(',') if origin.strip()]
    else:
        # Default origins for development (can be removed if not needed)
        allowed_origins = [
            "http://localhost:3000",
            "http://localhost:3001",
            "http://localhost:3002"
        ]
    
    CORS(
        app,
        origins=allowed_origins,
        supports_credentials=True,
        allow_headers=["Content-Type", "Authorization"],
        methods=["GET", "POST", "PUT", "DELETE", "OPTIONS"]
    )

    mongo.init_app(app) # Initialize the global mongo instance with the app instance
    app.mongo = mongo # Explicitly attach the mongo instance to the app object

    login_manager.init_app(app)
    login_manager.session_protection = "basic"  # Use basic session protection for cross-origin
    
    # Import blueprints and User model AFTER mongo and login_manager are initialized
    from routes.auth import auth_bp, User 
    from routes.projects import projects_bp
    app.register_blueprint(auth_bp)
    app.register_blueprint(projects_bp)
    
    # Add after_request handler to ensure CORS headers are always set
    @app.after_request
    def after_request(response):
        # Ensure Access-Control-Allow-Credentials is set for all responses
        if 'Access-Control-Allow-Credentials' not in response.headers:
            response.headers['Access-Control-Allow-Credentials'] = 'true'
        return response

    @login_manager.user_loader
    def load_user(user_id):
        # Access PyMongo via current_app.mongo.db
        user_doc = current_app.mongo.db.users.find_one({'_id': ObjectId(user_id)})
        if user_doc:
            return User(user_doc)
        return None

    @login_manager.unauthorized_handler
    def unauthorized_callback():
        return jsonify({'error': 'Unauthorized'}), 401

    @app.route('/', defaults={'path': ''})
    @app.route('/<path:path>')
    def serve(path):
        if path != "" and os.path.exists(app.static_folder + '/' + path):
            return send_from_directory(app.static_folder, path)
        else:
            return send_from_directory(app.static_folder, 'index.html')

    return app

if __name__ == '__main__':
    app = create_app()
    app.run(debug=True, host='0.0.0.0', port=5001)
