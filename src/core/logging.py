"""
Custom logging setup for the Word AI Agent
"""

import logging
import logging.handlers
import os
from pathlib import Path
from datetime import datetime
import queue

# Create logs directory if it doesn't exist
LOGS_DIR = Path('logs')
LOGS_DIR.mkdir(exist_ok=True)

# Custom QueueHandler for GUI logging
class QueueHandler(logging.Handler):
    """
    A handler class that sends log records to a queue
    It can be used from different threads
    """
    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        try:
            # Add log record to queue
            self.log_queue.put_nowait(record)
        except (KeyboardInterrupt, SystemExit):
            raise
        except Exception:
            self.handleError(record)

# Set up root logger
logger = logging.getLogger('word_ai_agent')
logger.setLevel(logging.DEBUG)

# Create file handler with rotation
file_handler = logging.handlers.RotatingFileHandler(
    LOGS_DIR / f'word_ai_agent_{datetime.now().strftime("%Y%m%d")}.log',
    maxBytes=10*1024*1024,  # 10MB
    backupCount=5
)
file_handler.setLevel(logging.DEBUG)

# Create console handler that doesn't show typing simulator debug messages
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)

# Create formatter and add it to the handlers
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# Add handlers to logger
logger.addHandler(file_handler)
logger.addHandler(console_handler)

# Set specific module loggers
typing_simulator_logger = logging.getLogger('word_ai_agent.src.utils.typing_simulator')
typing_simulator_logger.setLevel(logging.INFO)  # Only show INFO and above for typing simulator

def setup_logger(name):
    """Set up and return a logger instance"""
    child_logger = logger.getChild(name)
    return child_logger

def get_logger(name):
    """Get a logger instance"""
    return logging.getLogger(f'word_ai_agent.{name}')

def set_log_level(level):
    """Set the log level for console handler"""
    console_handler.setLevel(getattr(logging, level))

def log_exception(e: Exception, logger: logging.Logger) -> None:
    """Log an exception with full traceback"""
    logger.error(f"Exception occurred: {str(e)}", exc_info=True)

def log_performance(metric_name: str, value: float, unit: str, logger: logging.Logger) -> None:
    """Log performance metrics"""
    logger.info(f"PERFORMANCE - {metric_name}: {value} {unit}")
