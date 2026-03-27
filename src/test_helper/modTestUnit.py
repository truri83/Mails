"""
Module Name: Test Unit Helper

Purpose: Helper module for unit testing in the Mails project

Author: Mails Project
Date: 2026-03-27
"""

from typing import Any, List, Optional
from datetime import datetime

from src.models.email import Email


class TestEmailFactory:
    """Factory for creating test email objects."""

    @staticmethod
    def create_simple_email(
        sender: str = "test@example.com",
        recipient: str = "recipient@example.com",
        subject: str = "Test Email",
        body: str = "This is a test email."
    ) -> Email:
        """
        Create a simple test email.

        Args:
            sender: Sender email address
            recipient: Recipient email address
            subject: Email subject
            body: Email body

        Returns:
            Email object for testing
        """
        return Email(
            sender=sender,
            recipients=[recipient],
            subject=subject,
            body=body,
            timestamp=datetime(2026, 3, 27, 12, 0, 0)
        )

    @staticmethod
    def create_bulk_emails(count: int) -> List[Email]:
        """
        Create multiple test emails.

        Args:
            count: Number of emails to create

        Returns:
            List of test Email objects
        """
        emails = []
        for i in range(count):
            email = Email(
                sender=f"sender{i}@example.com",
                recipients=[f"recipient{i}@example.com"],
                subject=f"Test Email {i}",
                body=f"This is test email number {i}.",
                timestamp=datetime(2026, 3, 27, 12, i, 0)
            )
            emails.append(email)
        return emails


class MockEmailService:
    """Mock email service for testing."""

    def __init__(self):
        """Initialize the mock service."""
        self.sent_emails = []
        self.should_fail = False

    def send_email(self, email: Email) -> bool:
        """
        Mock send email method.

        Args:
            email: Email to send

        Returns:
            True if successful, False if configured to fail
        """
        if self.should_fail:
            return False

        self.sent_emails.append(email)
        return True

    def set_failure_mode(self, should_fail: bool):
        """
        Configure the mock to fail or succeed.

        Args:
            should_fail: Whether send operations should fail
        """
        self.should_fail = should_fail

    def reset(self):
        """Reset the mock service state."""
        self.sent_emails = []
        self.should_fail = False


def assert_email_valid(email: Email) -> bool:
    """
    Assert that an email object is valid.

    Args:
        email: Email to validate

    Returns:
        True if valid

    Raises:
        AssertionError: If email is invalid
    """
    assert email.sender, "Email must have a sender"
    assert email.recipients, "Email must have at least one recipient"
    assert email.subject, "Email must have a subject"
    assert email.body is not None, "Email must have a body"
    assert email.timestamp, "Email must have a timestamp"
    return True


def create_test_attachment(size: int = 1024) -> bytes:
    """
    Create test attachment data.

    Args:
        size: Size of the attachment in bytes

    Returns:
        Binary test data
    """
    return b'X' * size
