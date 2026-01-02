#!/usr/bin/env python3
"""
Email Sender Utility
Sends generated_excel files via email when DUMs are successfully processed
"""

import smtplib
import os
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import json

logger = logging.getLogger(__name__)

# Configuration file path
CONFIG_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'config', 'email_config.json')


def load_email_config():
    """Load email configuration from config file"""
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                return config
        return None
    except Exception as e:
        logger.error(f"Error loading email config: {e}")
        return None


def save_email_config(config):
    """Save email configuration to config file"""
    try:
        os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        logger.error(f"Error saving email config: {e}")
        return False


def send_excel_via_email(excel_file_path, lta_name, dum_number=None, recipient_email=None):
    """
    Send generated_excel file via email
    
    Args:
        excel_file_path: Path to the generated_excel*.xlsx file
        lta_name: Name of the LTA folder
        dum_number: Optional DUM number (if sending after single DUM)
        recipient_email: Optional recipient email (uses config if not provided)
    
    Returns:
        bool: True if email sent successfully, False otherwise
    """
    try:
        # Load configuration
        config = load_email_config()
        if not config:
            logger.warning("Email configuration not found. Email sending disabled.")
            return False
        
        # Get recipient email
        if not recipient_email:
            recipient_email = config.get('recipient_email')
        
        if not recipient_email:
            logger.error("No recipient email configured")
            return False
        
        # Email settings from config
        sender_email = config.get('sender_email')
        sender_password = config.get('sender_password')  # Gmail App Password
        smtp_server = config.get('smtp_server', 'smtp.gmail.com')
        smtp_port = config.get('smtp_port', 587)
        
        if not sender_email or not sender_password:
            logger.error("Email credentials not configured")
            return False
        
        # Check if file exists
        if not os.path.exists(excel_file_path):
            logger.error(f"Excel file not found: {excel_file_path}")
            return False
        
        # Create email message
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = f"✅ LTA Traité - {lta_name}"
        if dum_number:
            msg['Subject'] = f"✅ DUM {dum_number} Traité - {lta_name}"
        else:
            msg['Subject'] = f"✅ LTA Complet - {lta_name}"
        
        # Email body
        body = f"""
Bonjour,

Le fichier Excel généré pour le LTA "{lta_name}" est prêt.
"""
        if dum_number:
            body += f"\nDUM {dum_number} a été traité avec succès.\n"
        else:
            body += f"\nTous les DUMs de ce LTA ont été traités avec succès.\n"
        
        body += f"""
Fichier joint: {os.path.basename(excel_file_path)}
Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

Cordialement,
Système d'automatisation BADR
"""
        
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        # Attach Excel file
        with open(excel_file_path, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
        
        encoders.encode_base64(part)
        part.add_header(
            'Content-Disposition',
            f'attachment; filename= {os.path.basename(excel_file_path)}'
        )
        msg.attach(part)
        
        # Send email
        logger.info(f"Sending email to {recipient_email}...")
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        text = msg.as_string()
        server.sendmail(sender_email, recipient_email, text)
        server.quit()
        
        logger.info(f"✅ Email sent successfully to {recipient_email}")
        print(f"   📧 Email envoyé avec succès à {recipient_email}")
        return True
        
    except Exception as e:
        logger.error(f"Error sending email: {e}", exc_info=True)
        print(f"   ⚠️  Erreur envoi email: {e}")
        return False


def send_excel_after_dum_success(excel_file_path, lta_name, dum_number):
    """
    Convenience function to send Excel after DUM success
    """
    return send_excel_via_email(excel_file_path, lta_name, dum_number=dum_number)


def send_excel_after_lta_completion(excel_file_path, lta_name):
    """
    Convenience function to send Excel after LTA completion
    """
    return send_excel_via_email(excel_file_path, lta_name)
