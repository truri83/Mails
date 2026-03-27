"""
Module Name: Email Service

Purpose: Service layer for email operations

Author: Mails Project
Date: 2026-03-27
"""

import logging
from typing import List, Optional

from src.models.email import Email
from src.utils.email_utils import validate_email
from src.core.config import MAX_RETRIES, TIMEOUT

logger = logging.getLogger(__name__)


class EmailService:
    """Service for sending and managing emails."""

    def __init__(self, smtp_host: str, smtp_port: int):
        """
        Initialize EmailService.

        Args:
            smtp_host: SMTP server hostname
            smtp_port: SMTP server port
        """
        self.smtp_host = smtp_host
        self.smtp_port = smtp_port
        self.sent_emails = []

    def send_email(self, email: Email) -> bool:
        """
        Send an email.

        Args:
            email: Email object to send

        Returns:
            True if successful, False otherwise

        Raises:
            ValueError: If email addresses are invalid
        """
        # Validate sender
        if not validate_email(email.sender):
            raise ValueError(f"Invalid sender email: {email.sender}")

        # Validate recipients
        for recipient in email.recipients:
            if not validate_email(recipient):
                raise ValueError(f"Invalid recipient email: {recipient}")

        logger.info(f"Sending email from {email.sender} to {email.recipients}")

        try:
            # In a real implementation, this would connect to SMTP and send
            # For now, we just simulate success
            self.sent_emails.append(email)
            logger.debug(f"Email sent successfully: {email}")
            return True
        except Exception as e:
            logger.error(f"Failed to send email: {e}")
            return False

    def get_sent_count(self) -> int:
        """
        Get the number of sent emails.

        Returns:
            Count of sent emails
        """
        return len(self.sent_emails)

    def get_recent_emails(self, limit: int = 10) -> List[Email]:
        """
        Get recently sent emails.

        Args:
            limit: Maximum number of emails to return

        Returns:
            List of recent emails
        """
        return self.sent_emails[-limit:]
