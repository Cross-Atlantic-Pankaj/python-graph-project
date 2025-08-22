import os

class Config:
    # Flask settings
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'your-secret-key-here'
    
    # MongoDB settings
    MONGODB_SETTINGS = {
        'host': os.environ.get('MONGODB_URI') or 'mongodb://localhost:27017/graph_project'
    }
    
    # Matplotlib settings to reduce memory usage
    MATPLOTLIB_BACKEND = 'Agg'  # Non-interactive backend
    MATPLOTLIB_DPI = 150  # Reduced from 200 to save memory
    MATPLOTLIB_FIGSIZE = (10, 6)  # Standard figure size
    
    # File upload settings
    MAX_CONTENT_LENGTH = 100 * 1024 * 1024  # 100MB max file size
    UPLOAD_FOLDER = 'uploads'
    
    # Server performance settings
    THREADED = True
    PROCESSES = 1  # Single process to avoid matplotlib issues
    
    # Memory management settings
    GARBAGE_COLLECTION_INTERVAL = 5  # Force GC every 5 reports
    MAX_CHARTS_PER_REPORT = 50  # Limit charts per report
    
    # Temporary file settings
    TEMP_FILE_CLEANUP_INTERVAL = 300  # Clean temp files every 5 minutes
    
    # Logging settings
    LOG_LEVEL = 'INFO'
    
    # Development settings
    DEBUG = False
    TESTING = False

class DevelopmentConfig(Config):
    DEBUG = True
    LOG_LEVEL = 'DEBUG'

class ProductionConfig(Config):
    DEBUG = False
    LOG_LEVEL = 'WARNING'
    
    # Production optimizations
    MATPLOTLIB_DPI = 120  # Further reduced for production
    GARBAGE_COLLECTION_INTERVAL = 3  # More frequent GC in production

class TestingConfig(Config):
    TESTING = True
    DEBUG = True

config = {
    'development': DevelopmentConfig,
    'production': ProductionConfig,
    'testing': TestingConfig,
    'default': DevelopmentConfig
}
