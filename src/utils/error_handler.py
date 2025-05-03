"""
Error handling utility for the Word AI Agent
"""

import logging
import time
from typing import Dict, Any
from dataclasses import dataclass
from datetime import datetime
from abc import ABC, abstractmethod

from src.core.logging import get_logger

logger = get_logger(__name__)

@dataclass
class ErrorContext:
    """Context information for an error"""
    error_type: str
    timestamp: datetime
    message: str
    stack_trace: str
    retry_count: int = 0
    last_retry_time: datetime = None

class ErrorHandler:
    """Handles errors and retries with exponential backoff"""
    
    def __init__(self):
        self.error_counts: Dict[str, int] = {}
        self.max_retries = 3
        self.base_delay = 1.0
    
    def handle_error(self, error: Exception, context: Dict[str, Any] = None) -> None:
        """Handle an error with context"""
        try:
            error_type = type(error).__name__
            
            # Log the error
            logger.error(f"Error occurred: {str(error)}", exc_info=True)
            
            # Track error statistics
            self.error_counts[error_type] = self.error_counts.get(error_type, 0) + 1
            
            # Create error context
            error_context = ErrorContext(
                error_type=error_type,
                timestamp=datetime.now(),
                message=str(error),
                stack_trace=self._get_stack_trace(),
                retry_count=self.error_counts[error_type]
            )
            
            # Log error context
            logger.debug(f"Error context: {error_context}")
            
            # Attempt to recover
            if self._should_retry(error_context):
                self._retry_operation(error_context)
            else:
                self._escalate_error(error_context)
                
        except Exception as e:
            logger.critical(f"Error handling failed: {str(e)}")
            raise
    
    def _get_stack_trace(self) -> str:
        """Get the current stack trace"""
        import traceback
        return ''.join(traceback.format_stack())
    
    def _should_retry(self, context: ErrorContext) -> bool:
        """Determine if we should retry the operation"""
        return context.retry_count < self.max_retries
    
    def _retry_operation(self, context: ErrorContext) -> None:
        """Retry the operation with exponential backoff"""
        delay = self.base_delay * (2 ** context.retry_count)
        logger.info(f"Retrying operation in {delay:.1f} seconds...")
        time.sleep(delay)
    
    def _escalate_error(self, context: ErrorContext) -> None:
        """Escalate the error to a higher level"""
        logger.error(f"Error escalation: {context.error_type} occurred {context.retry_count} times")
        # Add error escalation logic here
    
    def get_error_statistics(self) -> Dict[str, int]:
        """Get statistics about errors"""
        return self.error_counts

class ErrorRecoveryStrategy:
    """Base class for error recovery strategies"""
    
    @abstractmethod
    def recover(self, error: Exception) -> bool:
        """Attempt to recover from an error"""
        pass

class RetryStrategy(ErrorRecoveryStrategy):
    """Retry-based error recovery strategy"""
    
    def __init__(self, max_retries: int = 3, base_delay: float = 1.0):
        self.max_retries = max_retries
        self.base_delay = base_delay
        self.retry_count = 0
    
    def recover(self, error: Exception) -> bool:
        """Attempt to recover by retrying"""
        if self.retry_count < self.max_retries:
            delay = self.base_delay * (2 ** self.retry_count)
            logger.info(f"Retrying operation in {delay:.1f} seconds...")
            time.sleep(delay)
            self.retry_count += 1
            return True
        return False
