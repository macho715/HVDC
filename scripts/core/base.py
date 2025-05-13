"""
Base Module
==========

This module contains the base class for all HVDC components.
"""

import logging
from pathlib import Path
from datetime import datetime
import json
import os

class HVDCBase:
    """Base class for all HVDC components."""
    
    def __init__(self, config_path: str = None):
        """
        Initialize the base class with common paths and logging setup.
        
        Args:
            config_path (str, optional): Path to the configuration file.
        """
        # Set up base paths
        self.root_dir = Path(__file__).parent.parent.parent
        self.data_dir = self.root_dir / 'data'
        self.output_dir = self.root_dir / 'output'
        self.log_dir = self.root_dir / 'logs'
        self.config_dir = self.root_dir / 'config'
        
        # Create necessary directories
        for dir_path in [self.data_dir, self.output_dir, self.log_dir, self.config_dir]:
            dir_path.mkdir(parents=True, exist_ok=True)
        
        # Load configuration
        self.config = self._load_config(config_path)
        
        # Set up logging
        self._setup_logging()
    
    def _load_config(self, config_path: str = None) -> dict:
        """
        Load configuration from file or use defaults.
        
        Args:
            config_path (str, optional): Path to the configuration file.
            
        Returns:
            dict: Configuration dictionary
        """
        default_config = {
            'log_level': 'INFO',
            'date_format': '%Y-%m-%d',
            'datetime_format': '%Y-%m-%d %H:%M:%S',
            'excel_template': 'template.xlsx',
            'output_prefix': 'HVDC'
        }
        
        if config_path and os.path.exists(config_path):
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    user_config = json.load(f)
                default_config.update(user_config)
            except Exception as e:
                self.logger.error(f"Error loading config file: {str(e)}")
        
        return default_config
    
    def _setup_logging(self):
        """Set up logging configuration."""
        log_file = self.log_dir / f"{self.__class__.__name__}.log"
        
        # Configure logging
        logging.basicConfig(
            level=getattr(logging, self.config['log_level']),
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        
        # Create logger for this class
        self.logger = logging.getLogger(self.__class__.__name__)
        self.logger.info(f"Initialized {self.__class__.__name__}")
    
    def get_timestamp(self) -> str:
        """
        Get current timestamp in configured format.
        
        Returns:
            str: Formatted timestamp
        """
        return datetime.now().strftime(self.config['datetime_format'])
    
    def get_date(self) -> str:
        """
        Get current date in configured format.
        
        Returns:
            str: Formatted date
        """
        return datetime.now().strftime(self.config['date_format'])
    
    def get_output_path(self, filename: str) -> Path:
        """
        Get full path for output file.
        
        Args:
            filename (str): Output filename
            
        Returns:
            Path: Full path to output file
        """
        return self.output_dir / filename
    
    def get_data_path(self, filename: str) -> Path:
        """
        Get full path for data file.
        
        Args:
            filename (str): Data filename
            
        Returns:
            Path: Full path to data file
        """
        return self.data_dir / filename
    
    def get_config_path(self, filename: str) -> Path:
        """
        Get full path for config file.
        
        Args:
            filename (str): Config filename
            
        Returns:
            Path: Full path to config file
        """
        return self.config_dir / filename 