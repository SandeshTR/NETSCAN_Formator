import os
import configparser
from logs.logs_handler import get_logger

logger = get_logger(__name__)


def get_creation_time(filename):
    return os.path.getctime(filename)


def load_config(config_path='devCode/config/config.ini'):
    """Load configuration from file"""
    logger.info(f'Loading configuration from file')
    config = configparser.ConfigParser()
    config.read(config_path)
    
    return {
        'unprocessed': config.get('general', 'unprocessed'),
        'input': config.get('general', 'inputpath'),
        'output': config.get('general', 'outputpath'),
        'error': config.get('general', 'errorpath'),
        'temp': config.get('general', 'temppath'),
        'process': config.get('general', 'processpath'),
        'archive': config.get('general', 'archive'),
        'root': config.get('general', 'rootpath')
    }