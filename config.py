"""
Configuration file for RPA Invoice Processing System
"""
import os
from pathlib import Path

# Base directories
BASE_DIR = Path(__file__).parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
PROCESSED_DIR = BASE_DIR / "processed"
ERROR_DIR = BASE_DIR / "error"
LOG_DIR = BASE_DIR / "logs"
TEMP_DIR = BASE_DIR / "temp"

# Create directories if they don't exist
for directory in [INPUT_DIR, OUTPUT_DIR, PROCESSED_DIR, ERROR_DIR, LOG_DIR, TEMP_DIR]:
    directory.mkdir(exist_ok=True)

# File extensions supported
SUPPORTED_EXTENSIONS = {
    'pdf': '.pdf',
    'excel': ['.xlsx', '.xls'],
    'word': ['.docx', '.doc']
}

# OCR Configuration
OCR_CONFIG = {
    'language': 'eng',
    'config': '--oem 3 --psm 6',
    'timeout': 30
}

# Database Configuration (if needed)
DATABASE_CONFIG = {
    'host': 'localhost',
    'port': 5432,
    'database': 'invoice_db',
    'user': 'admin',
    'password': 'password'
}

# Email Configuration (for Gmail integration)
EMAIL_CONFIG = {
    'smtp_server': 'smtp.gmail.com',
    'smtp_port': 587,
    'sender_email': 'gmail id',  # Replace with your Gmail address
    'sender_password': 'gmail app password'   # Replace with your Gmail app password
}

# Processing Configuration
PROCESSING_CONFIG = {
    'batch_size': 10,
    'max_retries': 3,
    'timeout': 300,  # 5 minutes
    'parallel_processing': True,
    'max_workers': 4
}

# Validation Rules
VALIDATION_RULES = {
    'required_fields': ['invoice_number', 'date', 'amount', 'vendor'],
    'amount_format': r'^\d+(\.\d{2})?$',
    'date_format': '%Y-%m-%d',
    'max_file_size': 10 * 1024 * 1024  # 10MB
}

# Logging Configuration
LOGGING_CONFIG = {
    'level': 'INFO',
    'format': '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    'file': LOG_DIR / 'rpa_system.log',
    'max_size': 10 * 1024 * 1024,  # 10MB
    'backup_count': 5
} 