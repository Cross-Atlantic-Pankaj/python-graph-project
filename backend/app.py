from flask import Flask, send_from_directory, jsonify, current_app
import re
import matplotlib.pyplot as plt
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
    # Load MONGO_URI from environment variable
    app.config["MONGO_URI"] = os.environ["MONGO_URI"]
    app.config['SECRET_KEY'] = 'your-secret-key-here'
    # Configure cookies to work over HTTP on AWS. When you move to HTTPS, set
    # SESSION_COOKIE_SECURE=True and SESSION_COOKIE_SAMESITE='None'.
    app.config['SESSION_COOKIE_SECURE'] = False
    app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
    CORS(
        app,
        origins=[
            "http://52.66.214.215:3000",
            "https://python-graph-project-fall-ss1cuw632.vercel.app",
            "https://python-graph-project-fall.vercel.app",
            "https://python-graph-project.onrender.com",
            "http://localhost:3000",
            "http://localhost:3001"
        ],
        supports_credentials=True,
        allow_headers=["Content-Type", "Authorization"],
        methods=["GET", "POST", "PUT", "DELETE", "OPTIONS"]
    )

    mongo.init_app(app) # Initialize the global mongo instance with the app instance
    app.mongo = mongo # Explicitly attach the mongo instance to the app object

    login_manager.init_app(app)
    
    # Import blueprints and User model AFTER mongo and login_manager are initialized
    from routes.auth import auth_bp, User 
    from routes.projects import projects_bp
    app.register_blueprint(auth_bp)
    app.register_blueprint(projects_bp)

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
