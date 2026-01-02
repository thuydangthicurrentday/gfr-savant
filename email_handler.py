"""
Email Handler for GOFILEROOM Downloader
Handles sending emails when critical errors occur
"""

import logging
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

logger = logging.getLogger(__name__)


class EmailHandler:
    """Class to handle sending emails when errors occur"""
    
    def __init__(self, smtp_server, smtp_port, sender_email, sender_password, recipient_emails, use_tls=True, enabled=True, machine_name=""):
        """
        Initialize EmailHandler
        
        Args:
            smtp_server (str): SMTP server (e.g., 'smtp.office365.com')
            smtp_port (int): SMTP port (e.g., 587)
            sender_email (str): Sender email
            sender_password (str): Sender email password
            recipient_emails (list): List of recipient emails
            use_tls (bool): Use TLS (default: True)
            enabled (bool): Whether to allow sending emails (default: True)
            machine_name (str): Machine name to include in email subject (default: "")
        """
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.sender_email = sender_email
        self.sender_password = sender_password
        self.recipient_emails = recipient_emails if isinstance(recipient_emails, list) else [recipient_emails]
        self.use_tls = use_tls
        self.enabled = enabled
        self.machine_name = machine_name or ""
    
    def send_error_email(self, subject, error_message, error_details=None):
        """
        Send error email
        
        Args:
            subject (str): Email subject
            error_message (str): Main error message
            error_details (dict): Error details (optional)
            
        Returns:
            bool: True if sent successfully, False if error or disabled
        """
        if not self.enabled:
            logger.info(f"Email sending is disabled. Would send: {subject}")
            return False
        
        try:
            # Add machine name to subject if available
            if self.machine_name:
                subject = f"[{self.machine_name}] {subject}"
            
            # Create message
            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = ', '.join(self.recipient_emails)
            msg['Subject'] = subject
            
            # Create email content
            body = f"""
            <html>
            <head></head>
            <body>
                <h2>GOFILEROOM Downloader - Critical Error</h2>
                <p><strong>Time:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
                <p><strong>Error Message:</strong></p>
                <p>{error_message}</p>
            """
            
            if error_details:
                body += "<p><strong>Error Details:</strong></p><ul>"
                for key, value in error_details.items():
                    body += f"<li><strong>{key}:</strong> {value}</li>"
                body += "</ul>"
            
            body += """
                <p><em>Please check log file for more details.</em></p>
            </body>
            </html>
            """
            
            msg.attach(MIMEText(body, 'html', 'utf-8'))
            
            # Send email
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                if self.use_tls:
                    server.starttls()
                server.login(self.sender_email, self.sender_password)
                server.send_message(msg)
            
            logger.info(f"Sent error email: {subject}")
            return True
            
        except Exception as e:
            logger.error(f"Error sending email: {str(e)}")
            return False
    
    def send_critical_error_email(self, consecutive_errors, error_summary):
        """
        Send email when critical error occurs (10 consecutive client errors)
        
        Args:
            consecutive_errors (int): Number of consecutive errors
            error_summary (dict): Summary of errors
            
        Returns:
            bool: True if sent successfully, False if error
        """
        subject = f"[CRITICAL] GOFILEROOM Downloader - {consecutive_errors} Consecutive Client Errors"
        error_message = f"""
        System encountered {consecutive_errors} consecutive errors while processing clients.
        Automation process has been stopped to avoid further errors.
        
        Please check:
        1. Network connection
        2. GOFILEROOM website status
        3. Log file for error details
        """
        
        return self.send_error_email(subject, error_message, error_summary)


def create_email_handler_from_config(config):
    """
    Create EmailHandler from config dictionary
    
    Args:
        config (dict): Dictionary containing email configuration from .env
        
    Returns:
        EmailHandler or None if no configuration or disabled
    """
    try:
        # Check if email is enabled
        enable_email = config.get('ENABLE_EMAIL', 'False').lower() in ('true', '1', 'yes')
        if not enable_email:
            logger.info("Email sending is disabled (ENABLE_EMAIL=False)")
            return None
        
        # Read config from .env with new variable names
        smtp_server = config.get('EMAIL_HOST')
        smtp_port = int(config.get('EMAIL_PORT', 587))
        sender_email = config.get('EMAIL_HOST_USER')
        sender_password = config.get('EMAIL_HOST_PASSWORD')
        default_from_email = config.get('DEFAULT_FROM_EMAIL', sender_email)
        recipient_list_str = config.get('EMAIL_recipient_list', '')
        use_tls = config.get('EMAIL_USE_TLS', 'True').lower() in ('true', '1', 'yes')
        machine_name = config.get('MACHINE', '')
        
        # Parse recipient list
        recipient_emails = [email.strip() for email in recipient_list_str.split(',') if email.strip()]
        
        if not all([smtp_server, sender_email, sender_password]):
            logger.warning("Insufficient email configuration (EMAIL_HOST, EMAIL_HOST_USER, EMAIL_HOST_PASSWORD), email sending will be disabled")
            return None
        
        if not recipient_emails:
            logger.warning("No recipient emails (EMAIL_recipient_list), email sending will be disabled")
            return None
        
        return EmailHandler(
            smtp_server=smtp_server,
            smtp_port=smtp_port,
            sender_email=default_from_email,  # Use DEFAULT_FROM_EMAIL as sender
            sender_password=sender_password,
            recipient_emails=recipient_emails,
            use_tls=use_tls,
            enabled=True,
            machine_name=machine_name
        )
        
    except Exception as e:
        logger.error(f"Error creating EmailHandler: {str(e)}")
        return None
