"""
Module Name: Configuration

Purpose: Central configuration for the Mails application

Author: Mails Project
Date: 2026-03-27
"""

import os

# Database configuration
DATABASE_HOST = os.getenv('DATABASE_HOST', 'localhost')
DATABASE_PORT = int(os.getenv('DATABASE_PORT', '5432'))
DATABASE_NAME = os.getenv('DATABASE_NAME', 'mails')

# Email server configuration
SMTP_HOST = os.getenv('SMTP_HOST', 'smtp.gmail.com')
SMTP_PORT = int(os.getenv('SMTP_PORT', '587'))
SMTP_USE_TLS = os.getenv('SMTP_USE_TLS', 'True').lower() == 'true'

# Application settings
DEBUG = os.getenv('DEBUG', 'False').lower() == 'true'
MAX_RETRIES = 3
TIMEOUT = 30

# Logging
LOG_LEVEL = os.getenv('LOG_LEVEL', 'INFO')
LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
