import os
import json
from flask import Blueprint, request, jsonify, current_app
from flask_login import login_required, current_user
from werkzeug.utils import secure_filename
from bson.objectid import ObjectId
from datetime import datetime 
import openpyxl
import tempfile
import re
import zipfile
import shutil

# Files are now stored in database, no upload folder needed

ALLOWED_EXTENSIONS = {'txt', 'pdf', 'png', 'jpg', 'jpeg', 'gif', 'csv', 'xlsx', 'docx'}
ALLOWED_REPORT_EXTENSIONS = {'csv', 'xlsx'}

projects_bp = Blueprint('projects', __name__)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def allowed_report_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_REPORT_EXTENSIONS

@projects_bp.route('/api/projects', methods=['GET'])
@login_required
def get_projects():
    """Get all projects for the current user"""
    try:
        projects = list(current_app.mongo.db.projects.find({'user_id': current_user.id}))
        
        # Convert ObjectId to string for JSON serialization
        for project in projects:
            project['_id'] = str(project['_id'])
            if 'created_at' in project:
                project['created_at'] = project['created_at'].isoformat()
            if 'updated_at' in project:
                project['updated_at'] = project['updated_at'].isoformat()
        
        return jsonify(projects)
    except Exception as e:
        current_app.logger.error(f"Error getting projects: {e}")
        return jsonify({'error': 'Failed to get projects'}), 500

@projects_bp.route('/api/projects', methods=['POST'])
@login_required
def create_project():
    """Create a new project"""
    try:
        data = request.get_json()
        
        if not data or 'name' not in data:
            return jsonify({'error': 'Project name is required'}), 400
        
        project = {
            'name': data['name'],
            'description': data.get('description', ''),
            'user_id': current_user.id,
            'created_at': datetime.utcnow(),
            'updated_at': datetime.utcnow()
        }
        
        result = current_app.mongo.db.projects.insert_one(project)
        project['_id'] = str(result.inserted_id)
        
        return jsonify(project), 201
    except Exception as e:
        current_app.logger.error(f"Error creating project: {e}")
        return jsonify({'error': 'Failed to create project'}), 500

@projects_bp.route('/api/projects/<project_id>', methods=['PUT'])
@login_required
def update_project(project_id):
    """Update a project"""
    try:
        data = request.get_json()
        
        if not data:
            return jsonify({'error': 'No data provided'}), 400
        
        # Verify project belongs to user
        project = current_app.mongo.db.projects.find_one({
            '_id': ObjectId(project_id),
            'user_id': current_user.id
        })
        
        if not project:
            return jsonify({'error': 'Project not found'}), 404
        
        # Update project
        update_data = {
            'updated_at': datetime.utcnow()
        }
        
        if 'name' in data:
            update_data['name'] = data['name']
        if 'description' in data:
            update_data['description'] = data['description']
        
        current_app.mongo.db.projects.update_one(
            {'_id': ObjectId(project_id)},
            {'$set': update_data}
        )
        
        return jsonify({'message': 'Project updated successfully'})
    except Exception as e:
        current_app.logger.error(f"Error updating project: {e}")
        return jsonify({'error': 'Failed to update project'}), 500

@projects_bp.route('/api/projects/<project_id>', methods=['DELETE'])
@login_required
def delete_project(project_id):
    """Delete a project"""
    try:
        # Verify project belongs to user
        project = current_app.mongo.db.projects.find_one({
            '_id': ObjectId(project_id),
            'user_id': current_user.id
        })
        
        if not project:
            return jsonify({'error': 'Project not found'}), 404
        
        # Delete project
        current_app.mongo.db.projects.delete_one({'_id': ObjectId(project_id)})
        
        return jsonify({'message': 'Project deleted successfully'})
    except Exception as e:
        current_app.logger.error(f"Error deleting project: {e}")
        return jsonify({'error': 'Failed to delete project'}), 500

@projects_bp.route('/api/chart', methods=['POST'])
@login_required
def create_chart():
    """Create a chart (simplified version for deployment)"""
    try:
        return jsonify({
            'error': 'Chart functionality is not available in this deployment. Please use the full version with matplotlib, pandas, and plotly dependencies.'
        }), 501
    except Exception as e:
        current_app.logger.error(f"Error creating chart: {e}")
        return jsonify({'error': 'Failed to create chart'}), 500 