#!/usr/bin/env python3
"""
Test script to debug the update project functionality
"""

import requests
import json

# Test configuration
BASE_URL = "http://52.66.214.215:5001"
PROJECT_ID = "688db341001db2a96607d894"  # From the error URL

def test_update_project_without_file():
    """Test updating project without file upload"""
    url = f"{BASE_URL}/api/projects/{PROJECT_ID}"
    
    data = {
        'name': 'test2_updated',
        'description': 'testing2_updated'
    }
    
    try:
        response = requests.put(url, data=data)
        print(f"Status Code: {response.status_code}")
        print(f"Response: {response.text}")
        return response
    except Exception as e:
        print(f"Error: {e}")
        return None

def test_update_project_with_file():
    """Test updating project with file upload"""
    url = f"{BASE_URL}/api/projects/{PROJECT_ID}"
    
    data = {
        'name': 'test2_with_file',
        'description': 'testing2_with_file'
    }
    
    # Create a small test file
    files = {
        'file': ('test.docx', b'This is a test file content', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    }
    
    try:
        response = requests.put(url, data=data, files=files)
        print(f"Status Code: {response.status_code}")
        print(f"Response: {response.text}")
        return response
    except Exception as e:
        print(f"Error: {e}")
        return None

if __name__ == "__main__":
    print("Testing update project without file...")
    test_update_project_without_file()
    
    print("\nTesting update project with file...")
    test_update_project_with_file()
