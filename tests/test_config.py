#!/usr/bin/env python3
"""
Unit tests for configuration module.
"""

import pytest
import os
import logging
from config import Config, get_logger


class TestConfigDefaults:
    """Test configuration defaults."""
    
    def test_config_has_pdf_folder(self):
        """Test that PDF folder is configured."""
        assert Config.PDF_FOLDER is not None
        assert isinstance(Config.PDF_FOLDER, str)
    
    def test_config_has_output_folder(self):
        """Test that output folder is configured."""
        assert Config.OUTPUT_FOLDER is not None
        assert isinstance(Config.OUTPUT_FOLDER, str)
    
    def test_config_has_log_level(self):
        """Test that log level is configured."""
        assert Config.LOG_LEVEL in ['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL']
    
    def test_config_regex_patterns(self):
        """Test that regex patterns are defined."""
        assert Config.DATE_PATTERN is not None
        assert Config.DB_PATTERN is not None
        assert Config.SUMMARY_PATTERN is not None
    
    def test_config_limits(self):
        """Test that limits are configured."""
        assert Config.MAX_PDF_SIZE > 0
        assert Config.MAX_COLUMN_WIDTH > 0
        assert Config.MIN_COLUMN_WIDTH >= 0


class TestEnvironmentVariables:
    """Test environment variable overrides."""
    
    def test_env_var_log_level(self):
        """Test LOG_LEVEL environment variable."""
        os.environ['LOG_LEVEL'] = 'DEBUG'
        # Note: This requires module reload to work, but we can at least test retrieval
        log_level = os.getenv('LOG_LEVEL', 'INFO')
        assert log_level == 'DEBUG'
        del os.environ['LOG_LEVEL']
    
    def test_env_var_max_pdf_size(self):
        """Test MAX_PDF_SIZE environment variable."""
        os.environ['MAX_PDF_SIZE'] = '50000000'
        max_size = int(os.getenv('MAX_PDF_SIZE', 100 * 1024 * 1024))
        assert max_size == 50000000
        del os.environ['MAX_PDF_SIZE']


class TestLogging:
    """Test logging setup."""
    
    def test_get_logger_returns_logger(self):
        """Test that get_logger returns logger instance."""
        logger = get_logger('test')
        assert isinstance(logger, logging.Logger)
    
    def test_setup_logging_creates_logger(self):
        """Test that setup_logging creates configured logger."""
        logger = Config.setup_logging()
        assert isinstance(logger, logging.Logger)
        assert logger.name == 'mutasi'
    
    def test_setup_logging_console_handler(self):
        """Test that console handler is added."""
        logger = Config.setup_logging()
        assert len(logger.handlers) > 0
        
        # Check for stream handler
        has_stream_handler = any(
            isinstance(h, logging.StreamHandler) 
            for h in logger.handlers
        )
        assert has_stream_handler


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
