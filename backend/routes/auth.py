from flask import Blueprint, request, jsonify, current_app
from flask_login import login_user, logout_user, login_required, current_user, UserMixin
from werkzeug.security import generate_password_hash, check_password_hash
from bson.objectid import ObjectId
# REMOVED: from app import mongo # Do NOT import mongo directly here

auth_bp = Blueprint('auth', __name__)

class User(UserMixin):
    def __init__(self, user_doc):
        self.id = str(user_doc['_id'])
        self.username = user_doc['username']
        self.full_name = user_doc['full_name']
        self.email = user_doc['email']
    def get_id(self):
        return self.id

@auth_bp.route('/api/register', methods=['POST'])
def register():
    data = request.get_json()
    if not data or not all(field in data for field in ['full_name', 'username', 'email', 'password']):
        return jsonify({'error': 'Missing required fields'}), 400
    
    # Access MongoDB via current_app.mongo.db
    if current_app.mongo.db.users.find_one({'username': data['username']}):
        return jsonify({'error': 'Username already exists'}), 400
    if current_app.mongo.db.users.find_one({'email': data['email']}):
        return jsonify({'error': 'Email already exists!'}), 400 
    
    user_id = current_app.mongo.db.users.insert_one({
        'full_name': data['full_name'],
        'username': data['username'],
        'email': data['email'],
        'password_hash': generate_password_hash(data['password'])
    }).inserted_id
    return jsonify({'message': 'Registration successful'}), 201

@auth_bp.route('/api/login', methods=['POST'])
def login():
    data = request.get_json()
    if not data or not data.get('username') or not data.get('password'):
        return jsonify({'error': 'Missing username or password'}), 400
    # Access MongoDB via current_app.mongo.db
    user_doc = current_app.mongo.db.users.find_one({'username': data['username']})
    if user_doc and check_password_hash(user_doc['password_hash'], data['password']):
        user = User(user_doc)
        login_user(user)
        return jsonify({'message': 'Login successful', 'user': {
            'id': user.id,
            'username': user.username,
            'full_name': user.full_name,
            'email': user.email
        }}), 200
    return jsonify({'error': 'Invalid username or password'}), 401

@auth_bp.route('/api/user')
@login_required
def get_user():
    # Access MongoDB via current_app.mongo.db
    user_doc = current_app.mongo.db.users.find_one({'_id': ObjectId(current_user.get_id())})
    if not user_doc:
        return jsonify({'error': 'User not found'}), 404
    return jsonify({'user': {
        'id': str(user_doc['_id']),
        'username': user_doc['username'],
        'full_name': user_doc['full_name'],
        'email': user_doc['email']
    }}), 200

@auth_bp.route('/api/logout')
@login_required
def logout():
    logout_user()
    return jsonify({'message': 'Logged out successfully'}), 200
