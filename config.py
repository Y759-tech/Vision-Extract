import os
from datetime import timedelta

class Config:
    """Configuration de l'application"""
    
    # Database
    SQL_SERVER = os.getenv('SQL_SERVER', 'localhost\\SQLEXPRESS')
    DATABASE = os.getenv('DATABASE', 'REMUCI_VISION')
    
    # Security
    SECRET_KEY = os.getenv('SECRET_KEY', 'votre-cle-secrete-changez-moi')
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB max file size
    
    # API Limits
    RATE_LIMIT = os.getenv('RATE_LIMIT', '100 per hour')
    
    # CORS
    CORS_ORIGINS = os.getenv('CORS_ORIGINS', 'http://localhost:5000,http://127.0.0.1:5000')

class DevelopmentConfig(Config):
    DEBUG = True
    TESTING = False

class ProductionConfig(Config):
    DEBUG = False
    TESTING = False

config = {
    'development': DevelopmentConfig,
    'production': ProductionConfig,
    'default': DevelopmentConfig
}