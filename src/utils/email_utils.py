"""
Module Name: Email Utilities

Purpose: Utility functions for email processing

Author: Mails Project
Date: 2026-03-27
"""

import re
from typing import Optional


def validate_email(email: str) -> bool:
    """
    Validate an email address format.

    Args:
        email: Email address to validate

    Returns:
        True if email is valid, False otherwise
    """
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, email))


def parse_email_header(header: str) -> dict:
    """
    Parse email header into components.

    Args:
        header: Raw email header string

    Returns:
        Dictionary with parsed header components
    """
    components = {}
    lines = header.split('\n')

    for line in lines:
        if ':' in line:
            key, value = line.split(':', 1)
            components[key.strip()] = value.strip()

    return components


def format_email_address(name: str, email: str) -> str:
    """
    Format a name and email into standard format.

    Args:
        name: Display name
        email: Email address

    Returns:
        Formatted email address string
    """
    if name:
        return f'"{name}" <{email}>'
    return email
