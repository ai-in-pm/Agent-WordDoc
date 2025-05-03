"""
Security module for the Word AI Agent

Provides input validation, credential management, and security features.
"""

import re
import os
import json
import hashlib
import base64
import secrets
import time
from typing import Dict, Any, List, Optional, Union
from pathlib import Path
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

from src.core.logging import get_logger

logger = get_logger(__name__)

class InputValidator:
    """Validates user input to prevent security issues"""
    
    @staticmethod
    def validate_filename(filename: str) -> bool:
        """Validate a filename for security"""
        # Check for directory traversal attacks
        if '..' in filename or '/' in filename or '\\' in filename:
            logger.warning(f"Invalid filename: {filename}")
            return False
        
        # Check for invalid characters
        if re.search(r'[<>:"|?*]', filename):
            logger.warning(f"Invalid characters in filename: {filename}")
            return False
        
        return True
    
    @staticmethod
    def validate_path(path: str) -> bool:
        """Validate a path for security"""
        try:
            # Convert to absolute path and check if it exists
            path = os.path.abspath(path)
            return os.path.exists(path)
        except Exception as e:
            logger.warning(f"Invalid path: {path} - {str(e)}")
            return False
    
    @staticmethod
    def sanitize_input(input_string: str) -> str:
        """Sanitize user input for security"""
        # Remove potentially dangerous characters
        sanitized = re.sub(r'[\\\'"<>]', '', input_string)
        return sanitized
    
    @staticmethod
    def validate_api_key(api_key: str) -> bool:
        """Validate an API key format"""
        # Check for common API key patterns
        # This is a simple example - real validation would depend on the API
        if not api_key or len(api_key) < 10:
            logger.warning("API key is too short or empty")
            return False
        
        return True

class CredentialManager:
    """Securely manages credentials and API keys"""
    
    def __init__(self, encryption_key: Optional[str] = None):
        self.credentials_file = Path(".credentials")
        self.encryption_key = self._get_encryption_key(encryption_key)
        self.fernet = Fernet(self.encryption_key)
    
    def _get_encryption_key(self, provided_key: Optional[str]) -> bytes:
        """Get or generate an encryption key"""
        if provided_key:
            # Use the provided key
            salt = b'8xTu9c3K2j5G1z4A'  # Fixed salt for key derivation
            kdf = PBKDF2HMAC(
                algorithm=hashes.SHA256(),
                length=32,
                salt=salt,
                iterations=100000
            )
            return base64.urlsafe_b64encode(kdf.derive(provided_key.encode()))
        else:
            # Generate a new key
            key_file = Path(".encryption_key")
            if key_file.exists():
                return key_file.read_bytes()
            else:
                key = Fernet.generate_key()
                key_file.write_bytes(key)
                return key
    
    def store_credential(self, name: str, value: str) -> None:
        """Securely store a credential"""
        encrypted_value = self.fernet.encrypt(value.encode()).decode()
        
        credentials = {}
        if self.credentials_file.exists():
            try:
                encrypted_data = self.credentials_file.read_bytes()
                decrypted_data = self.fernet.decrypt(encrypted_data).decode()
                credentials = json.loads(decrypted_data)
            except Exception as e:
                logger.error(f"Error reading credentials: {str(e)}")
        
        credentials[name] = encrypted_value
        
        encrypted_json = self.fernet.encrypt(json.dumps(credentials).encode())
        self.credentials_file.write_bytes(encrypted_json)
        
        logger.info(f"Stored credential: {name}")
    
    def get_credential(self, name: str) -> Optional[str]:
        """Securely retrieve a credential"""
        if not self.credentials_file.exists():
            logger.warning("Credentials file does not exist")
            return None
        
        try:
            encrypted_data = self.credentials_file.read_bytes()
            decrypted_data = self.fernet.decrypt(encrypted_data).decode()
            credentials = json.loads(decrypted_data)
            
            if name not in credentials:
                logger.warning(f"Credential {name} not found")
                return None
            
            encrypted_value = credentials[name]
            return self.fernet.decrypt(encrypted_value.encode()).decode()
        except Exception as e:
            logger.error(f"Error retrieving credential {name}: {str(e)}")
            return None
    
    def delete_credential(self, name: str) -> bool:
        """Delete a stored credential"""
        if not self.credentials_file.exists():
            return False
        
        try:
            encrypted_data = self.credentials_file.read_bytes()
            decrypted_data = self.fernet.decrypt(encrypted_data).decode()
            credentials = json.loads(decrypted_data)
            
            if name not in credentials:
                return False
            
            del credentials[name]
            
            encrypted_json = self.fernet.encrypt(json.dumps(credentials).encode())
            self.credentials_file.write_bytes(encrypted_json)
            
            logger.info(f"Deleted credential: {name}")
            return True
        except Exception as e:
            logger.error(f"Error deleting credential {name}: {str(e)}")
            return False

class RateLimiter:
    """Implements rate limiting to prevent API abuse"""
    
    def __init__(self, max_requests: int = 60, time_window: int = 60):
        self.max_requests = max_requests  # Max requests per time window
        self.time_window = time_window  # Time window in seconds
        self.request_history = []
    
    def is_rate_limited(self) -> bool:
        """Check if the current request should be rate limited"""
        current_time = time.time()
        
        # Remove old requests from history
        self.request_history = [t for t in self.request_history if current_time - t < self.time_window]
        
        # Check if we've exceeded the rate limit
        if len(self.request_history) >= self.max_requests:
            logger.warning("Rate limit exceeded")
            return True
        
        # Add the current request to history
        self.request_history.append(current_time)
        return False
    
    def get_remaining_requests(self) -> int:
        """Get the number of remaining requests in the current time window"""
        current_time = time.time()
        
        # Remove old requests from history
        self.request_history = [t for t in self.request_history if current_time - t < self.time_window]
        
        return max(0, self.max_requests - len(self.request_history))
    
    def get_reset_time(self) -> float:
        """Get the time until the rate limit resets"""
        if not self.request_history:
            return 0.0
        
        current_time = time.time()
        oldest_request = min(self.request_history)
        
        return max(0.0, self.time_window - (current_time - oldest_request))

# Added a proper class closing
