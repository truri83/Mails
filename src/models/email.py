"""
Module Name: Email Model

Purpose: Data model for email objects

Author: Mails Project
Date: 2026-03-27
"""

from datetime import datetime
from typing import List, Optional


class Email:
    """Represents an email message."""

    def __init__(
        self,
        sender: str,
        recipients: List[str],
        subject: str,
        body: str,
        timestamp: Optional[datetime] = None
    ):
        """
        Initialize an Email object.

        Args:
            sender: Email address of the sender
            recipients: List of recipient email addresses
            subject: Email subject line
            body: Email body content
            timestamp: When the email was sent (defaults to now)
        """
        self.sender = sender
        self.recipients = recipients
        self.subject = subject
        self.body = body
        self.timestamp = timestamp or datetime.now()
        self.attachments = []

    def add_attachment(self, filename: str, content: bytes):
        """
        Add an attachment to the email.

        Args:
            filename: Name of the attachment file
            content: Binary content of the attachment
        """
        self.attachments.append({
            'filename': filename,
            'content': content
        })

    def __repr__(self) -> str:
        """String representation of the email."""
        return f"Email(from={self.sender}, to={self.recipients}, subject='{self.subject}')"

    def to_dict(self) -> dict:
        """
        Convert email to dictionary format.

        Returns:
            Dictionary representation of the email
        """
        return {
            'sender': self.sender,
            'recipients': self.recipients,
            'subject': self.subject,
            'body': self.body,
            'timestamp': self.timestamp.isoformat(),
            'attachments': len(self.attachments)
        }
