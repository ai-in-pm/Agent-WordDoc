"""
Error Logger for Word Interface Explorer

Captures and logs all failure attempts by the AI Agent when interacting with Microsoft Word.
"""

import os
import sys
import logging
import traceback
import datetime
from pathlib import Path
import json

# Define the log directory
LOG_DIR = Path(r"C:\Users\djjme\OneDrive\Desktop\CC-Directory\Agent WordDoc\src\bootstrap_training\word_interface\i_failed_logs")

# Ensure log directory exists
LOG_DIR.mkdir(parents=True, exist_ok=True)

class ErrorLogger:
    """Error logger for Word Interface Explorer"""
    
    def __init__(self, module_name="word_interface"):
        """Initialize the error logger"""
        self.module_name = module_name
        self.log_file = None
        self.setup_logger()
    
    def setup_logger(self):
        """Set up the logger"""
        # Create a timestamped log file name
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_file = LOG_DIR / f"failure_{self.module_name}_{timestamp}.log"
        
        # Configure file handler
        file_handler = logging.FileHandler(self.log_file)
        file_handler.setLevel(logging.DEBUG)
        
        # Configure formatter
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        file_handler.setFormatter(formatter)
        
        # Get logger and add handler
        self.logger = logging.getLogger(self.module_name)
        self.logger.setLevel(logging.DEBUG)
        self.logger.addHandler(file_handler)
    
    def log_error(self, error, context=None):
        """Log an error with optional context"""
        # Get full traceback
        tb = traceback.format_exc()
        
        # Log the error
        self.logger.error(f"Error: {str(error)}")
        self.logger.error(f"Traceback: {tb}")
        
        # Log context if provided
        if context:
            context_str = json.dumps(context, indent=2, default=str)
            self.logger.error(f"Context: {context_str}")
        
        # Return the path to the log file
        return self.log_file
    
    def log_failure(self, title, message, context=None):
        """Log a failure with a title and message"""
        self.logger.error(f"Failure: {title}")
        self.logger.error(f"Message: {message}")
        
        # Log context if provided
        if context:
            context_str = json.dumps(context, indent=2, default=str)
            self.logger.error(f"Context: {context_str}")
        
        # Return the path to the log file
        return self.log_file
    
    def log_error_screenshot(self, error, screenshot_path):
        """Log an error with a screenshot"""
        self.logger.error(f"Error with screenshot: {str(error)}")
        self.logger.error(f"Screenshot: {screenshot_path}")
        
        # Return the path to the log file
        return self.log_file
    
    def get_all_logs(self):
        """Get a list of all log files"""
        return list(LOG_DIR.glob("*.log"))

# Create a global error logger instance
error_logger = ErrorLogger()

def log_error(error, context=None):
    """Convenience function to log an error"""
    return error_logger.log_error(error, context)

def log_failure(title, message, context=None):
    """Convenience function to log a failure"""
    return error_logger.log_failure(title, message, context)

def get_all_logs():
    """Convenience function to get all logs"""
    return error_logger.get_all_logs()
