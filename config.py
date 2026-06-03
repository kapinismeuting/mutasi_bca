#!/usr/bin/env python3
"""
Configuration management for BCA bank statement extractor.
Supports environment variables and default values.
"""

import os
import logging
from pathlib import Path


class Config:
    """Application configuration with environment variable support."""
    
    # Input/Output paths
    PDF_FOLDER: str = os.getenv('PDF_FOLDER', os.path.expanduser('~/dev/appdev/Mutasi/2016'))
    OUTPUT_FOLDER: str = os.getenv('OUTPUT_FOLDER', os.path.expanduser('~/dev/appdev/Mutasi_Excel'))
    
    # Logging configuration
    LOG_LEVEL: str = os.getenv('LOG_LEVEL', 'INFO')
    LOG_FORMAT: str = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    LOG_FILE: str = os.getenv('LOG_FILE', None)
    
    # Processing configuration
    MAX_PDF_SIZE: int = int(os.getenv('MAX_PDF_SIZE', 100 * 1024 * 1024))  # 100 MB
    BACKUP_FILES: bool = os.getenv('BACKUP_FILES', 'true').lower() == 'true'
    
    # Regex patterns
    DATE_PATTERN: str = r'^(\d{2})/(\d{2})'
    DB_PATTERN: str = r'(\d{1,3}(?:,\d{3})*\.\d{2})\s*DB'
    CR_PATTERN: str = r'(\d{1,3}(?:,\d{3})*\.\d{2})'
    SALDO_PATTERN: str = r'(\d{1,3}(?:,\d{3})*\.\d{2})\s*$'
    SUMMARY_PATTERN: str = r'^(SALDO AWAL|MUTASI CR|MUTASI DB|SALDO AKHIR)\s*:'
    
    # Excel formatting
    MIN_COLUMN_WIDTH: int = 2
    MAX_COLUMN_WIDTH: int = 80
    DATE_LENGTH: int = 5  # DD/MM format
    SHEET_NAME_LENGTH: int = 3  # First 3 chars for month abbreviation
    
    @classmethod
    def validate_paths(cls) -> None:
        """Validate that configured paths are accessible."""
        pdf_path = Path(cls.PDF_FOLDER).resolve()
        output_path = Path(cls.OUTPUT_FOLDER).resolve()
        
        # Check if PDF folder exists
        if not pdf_path.exists():
            raise ValueError(f"PDF folder does not exist: {cls.PDF_FOLDER}")
        
        if not pdf_path.is_dir():
            raise ValueError(f"PDF folder is not a directory: {cls.PDF_FOLDER}")
    
    @classmethod
    def setup_logging(cls) -> logging.Logger:
        """Configure logging for the application."""
        logger = logging.getLogger('mutasi')
        
        # Remove existing handlers to avoid duplicates
        logger.handlers = []
        
        # Set log level
        log_level = getattr(logging, cls.LOG_LEVEL.upper(), logging.INFO)
        logger.setLevel(log_level)
        
        # Create formatter
        formatter = logging.Formatter(cls.LOG_FORMAT)
        
        # Console handler
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)
        
        # File handler (if configured)
        if cls.LOG_FILE:
            file_handler = logging.FileHandler(cls.LOG_FILE)
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
        
        return logger


def get_logger(name: str) -> logging.Logger:
    """Get a configured logger instance."""
    return logging.getLogger(name)
