"""
Configuration Management for Voice-Controlled PowerPoint System

This module handles loading configuration from environment variables
and .env files using python-dotenv and pydantic for validation.
"""

import os
import logging
from pathlib import Path
from typing import Optional
from pydantic import BaseModel, field_validator
from pydantic_settings import BaseSettings
from dotenv import load_dotenv


class AppConfig(BaseSettings):
    """Application configuration with validation."""
    
    # Gemini AI Configuration
    gemini_api_key: str = ""
    
    # Voice Recognition Settings
    voice_recognition_engine: str = "google"
    voice_language: str = "en-US"
    enable_voice_feedback: bool = True
    
    # Logging Configuration
    log_level: str = "INFO"
    log_format: str = "colored"
    log_file: str = "presentation_app.log"
    
    # Application Settings
    auto_save: bool = True
    default_presentation_name: str = "voice_presentation"
    max_slide_count: int = 50
    
    # AI Content Generation Settings
    ai_temperature: float = 0.7
    ai_max_tokens: int = 1000
    ai_content_language: str = "english"
    
    class Config:
        env_file = ".env"
        env_file_encoding = "utf-8"
        case_sensitive = False
    
    @field_validator('log_level')
    @classmethod
    def validate_log_level(cls, v):
        """Validate log level."""
        valid_levels = ['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL']
        if v.upper() not in valid_levels:
            raise ValueError(f'Log level must be one of: {valid_levels}')
        return v.upper()
    
    @field_validator('voice_recognition_engine')
    @classmethod
    def validate_voice_engine(cls, v):
        """Validate voice recognition engine."""
        valid_engines = ['google', 'sphinx', 'bing', 'azure']
        if v.lower() not in valid_engines:
            raise ValueError(f'Voice engine must be one of: {valid_engines}')
        return v.lower()
    
    @field_validator('ai_temperature')
    @classmethod
    def validate_temperature(cls, v):
        """Validate AI temperature parameter."""
        if not 0.0 <= v <= 2.0:
            raise ValueError('AI temperature must be between 0.0 and 2.0')
        return v


def load_config() -> AppConfig:
    """
    Load configuration from environment variables and .env file.
    
    Returns:
        AppConfig: Validated configuration object
    """
    # Look for .env file in current directory
    env_path = Path('.env')
    
    if env_path.exists():
        load_dotenv(env_path)
        print(f"Loaded configuration from {env_path}")
    else:
        # Look for .env.template as fallback
        template_path = Path('.env.template')
        if template_path.exists():
            print(f"No .env file found. Please copy {template_path} to .env and configure your settings.")
        else:
            print("No .env file found. Using default configuration.")
    
    try:
        config = AppConfig()
        return config
    except Exception as e:
        print(f"Configuration error: {e}")
        print("Using default configuration with empty API key.")
        # Return config with minimal settings if validation fails
        return AppConfig(gemini_api_key="")


def setup_logging(config: AppConfig) -> logging.Logger:
    """
    Setup logging configuration based on config settings.
    
    Args:
        config (AppConfig): Application configuration
        
    Returns:
        logging.Logger: Configured logger instance
    """
    # Create logs directory if it doesn't exist
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)
    
    # Setup log file path
    log_file_path = log_dir / config.log_file
    
    # Configure logging level
    log_level = getattr(logging, config.log_level)
    
    # Create logger
    logger = logging.getLogger("voice_ppt")
    logger.setLevel(log_level)
    
    # Remove existing handlers to avoid duplicates
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)
    
    # Create formatters
    if config.log_format == "colored":
        try:
            import colorlog
            console_formatter = colorlog.ColoredFormatter(
                "%(log_color)s%(asctime)s - %(name)s - %(levelname)s - %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S",
                log_colors={
                    'DEBUG': 'cyan',
                    'INFO': 'green',
                    'WARNING': 'yellow',
                    'ERROR': 'red',
                    'CRITICAL': 'red,bg_white',
                }
            )
        except ImportError:
            console_formatter = logging.Formatter(
                "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S"
            )
    else:
        console_formatter = logging.Formatter(
            "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S"
        )
    
    file_formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    
    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)
    console_handler.setFormatter(console_formatter)
    logger.addHandler(console_handler)
    
    # File handler
    file_handler = logging.FileHandler(log_file_path)
    file_handler.setLevel(logging.DEBUG)  # Always log DEBUG to file
    file_handler.setFormatter(file_formatter)
    logger.addHandler(file_handler)
    
    logger.info(f"Logging initialized - Level: {config.log_level}, File: {log_file_path}")
    
    return logger


def validate_gemini_api_key(api_key: str) -> bool:
    """
    Validate Gemini API key format.
    
    Args:
        api_key (str): API key to validate
        
    Returns:
        bool: True if API key appears valid
    """
    if not api_key:
        return False
    
    # Basic validation - Gemini API keys typically start with certain patterns
    if len(api_key) < 20:
        return False
    
    # Additional validation could be added here
    return True


# Global configuration instance
_config: Optional[AppConfig] = None
_logger: Optional[logging.Logger] = None


def get_config() -> AppConfig:
    """Get the global configuration instance."""
    global _config
    if _config is None:
        _config = load_config()
    return _config


def get_logger() -> logging.Logger:
    """Get the global logger instance."""
    global _logger
    if _logger is None:
        config = get_config()
        _logger = setup_logging(config)
    return _logger