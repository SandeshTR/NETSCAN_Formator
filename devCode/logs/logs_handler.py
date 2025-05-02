import os
import logging
from datetime import datetime
from typing import Optional, Dict, Union

import logging.handlers

# Default configurations
DEFAULT_LOG_LEVEL = logging.INFO
DEFAULT_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
DEFAULT_DATE_FORMAT = '%Y-%m-%d %H:%M:%S'
DEFAULT_LOG_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'logs')
DEFAULT_MAX_BYTES = 10 * 1024 * 1024  # 10 MB
DEFAULT_BACKUP_COUNT = 5

# Ensure log directory exists
os.makedirs(DEFAULT_LOG_DIR, exist_ok=True)


class LogsHandler:
    """
    Central logging handler for the application.
    All logging configuration and instantiation should go through this class.
    """
    
    _loggers: Dict[str, logging.Logger] = {}
    _initialized = False
    
    @classmethod
    def configure_logging(cls, 
                         log_level: int = DEFAULT_LOG_LEVEL,
                         log_format: str = DEFAULT_FORMAT,
                         date_format: str = DEFAULT_DATE_FORMAT,
                         log_dir: str = DEFAULT_LOG_DIR,
                         file_logging: bool = True,
                         console_logging: bool = True,
                         max_bytes: int = DEFAULT_MAX_BYTES,
                         backup_count: int = DEFAULT_BACKUP_COUNT) -> None:
        """
        Configure the root logger with the specified parameters.
        
        Args:
            log_level: The logging level (e.g., logging.INFO, logging.DEBUG)
            log_format: The format string for the log messages
            date_format: The format string for timestamps in log messages
            log_dir: Directory where log files will be stored
            file_logging: Whether to enable logging to file
            console_logging: Whether to enable logging to console
            max_bytes: Maximum size of each log file before rotation
            backup_count: Number of backup log files to keep
        """
        # Create formatter
        formatter = logging.Formatter(log_format, date_format)
        
        # Configure root logger
        root_logger = logging.getLogger()
        root_logger.setLevel(log_level)
        
        # Remove existing handlers to avoid duplicates
        for handler in list(root_logger.handlers):
            root_logger.removeHandler(handler)
        
        # Add console handler if requested
        if console_logging:
            console_handler = logging.StreamHandler()
            console_handler.setFormatter(formatter)
            root_logger.addHandler(console_handler)
        
        # Add file handler if requested
        if file_logging:
            os.makedirs(log_dir, exist_ok=True)
            log_file = os.path.join(log_dir, f"app_{datetime.now().strftime('%Y%m%d')}.log")
            file_handler = logging.handlers.RotatingFileHandler(
                log_file, maxBytes=max_bytes, backupCount=backup_count
            )
            file_handler.setFormatter(formatter)
            root_logger.addHandler(file_handler)
        
        cls._initialized = True
        
        # Log configuration info
        root_logger.info(f"Logging configured with level={logging.getLevelName(log_level)}")
        if file_logging:
            root_logger.info(f"File logging enabled at: {log_file}")

    @classmethod
    def get_logger(cls, name: str, 
                  log_level: Optional[int] = None) -> logging.Logger:
        """
        Get a named logger. If the logger hasn't been created yet,
        it will be created and stored for future use.
        
        Args:
            name: The name of the logger (typically __name__ from the calling module)
            log_level: Optional specific log level for this logger
            
        Returns:
            logging.Logger: The configured logger
        """
        # Ensure logging is configured
        if not cls._initialized:
            cls.configure_logging()
        
        # Create or retrieve the named logger
        if name not in cls._loggers:
            logger = logging.getLogger(name)
            if log_level is not None:
                logger.setLevel(log_level)
            cls._loggers[name] = logger
        
        return cls._loggers[name]

    @staticmethod
    def set_log_level(logger_name: Optional[str] = None, 
                     level: Union[int, str] = logging.INFO) -> None:
        """
        Change the log level of a specific logger or the root logger.
        
        Args:
            logger_name: Name of the logger to modify (None for root logger)
            level: New log level (can be integer or string like "INFO", "DEBUG")
        """
        # Convert string level to int if needed
        if isinstance(level, str):
            level = getattr(logging, level.upper())
        
        logger = logging.getLogger(logger_name) if logger_name else logging.getLogger()
        logger.setLevel(level)
        logger.info(f"Log level for {'root' if not logger_name else logger_name} changed to {logging.getLevelName(level)}")


# Convenience functions
def get_logger(name: str, log_level: Optional[int] = None) -> logging.Logger:
    """Get a configured logger with the specified name"""
    return LogsHandler.get_logger(name, log_level)

def configure_logging(**kwargs) -> None:
    """Configure the logging system with the specified parameters"""
    LogsHandler.configure_logging(**kwargs)

def set_log_level(logger_name: Optional[str] = None, level: Union[int, str] = logging.INFO) -> None:
    """Set log level for a specific logger or the root logger"""
    LogsHandler.set_log_level(logger_name, level)


# Example usage (commented out)
if __name__ == "__main__":
    # Configure logging
    configure_logging(log_level=logging.DEBUG)
    
    # Get a logger
    logger = get_logger(__name__)
    
    # # Log some messages
    # logger.debug("This is a debug message")
    # logger.info("This is an info message")
    # logger.warning("This is a warning message")
    # logger.error("This is an error message")
    
    # # Change log level
    # set_log_level(level=logging.WARNING)
    # logger.debug("This debug message won't appear")
    # logger.warning("But this warning will")