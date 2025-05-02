import os
import logging
import configparser
import shutil
from pathlib import Path
from datetime import datetime
from devCode import extract_input


def delete_all_files(folders):
    '''Delete all files in the specified folders.'''
    for folder in folders:
        logging.info(f'Processing folder: {folder}')
        
        # Check if folder exists
        if not os.path.exists(folder):
            logging.warning(f"Folder does not exist: {folder}")
            continue
            
        for filename in os.listdir(folder):
            file_path = os.path.join(folder, filename)

            # Check if it is a file before deleting
            if os.path.isfile(file_path):
                try: 
                    os.remove(file_path)
                    logging.info(f"Deleted file: {file_path}")
                except PermissionError:
                    logging.error(f"Permission denied when deleting {file_path}")
                except Exception as e:
                    logging.error(f"Failed to delete file {file_path}: {e}")


def get_creation_time(filename):
    return os.path.getctime(filename)


def get_sorted_zip_files(directory):
    """Get sorted zip files from directory based on creation time"""
    files = os.listdir(directory)
    sorted_files = sorted(
        (file for file in files if file.lower().endswith('.zip')),
        key=lambda f: get_creation_time(Path(os.path.join(directory, f)))
    )
    
    if not sorted_files:
        logging.warning("No .zip files present in input folder")
        
    return sorted_files


def process_zip_file(filename, config_paths):
    """Process a single zip file"""
    input_path = config_paths['input']
    output_path = config_paths['output']
    error_path = config_paths['error']
    temp_path = config_paths['temp']
    process_path = config_paths['process']
    archive_path = config_paths['archive']
    unprocessed_path = config_paths['unprocessed']
    
    sub_output_folder = datetime.now().strftime("%B_%d_%Y").upper()
    
    try:            
        logging.info(f'** Started Processing file: {filename} **')
        delete_all_files([temp_path, process_path]) 

        logging.info("File extraction started from input path")
        extract_input.Input_Extract(
            os.path.join(input_path, filename),
            output_path,
            error_path,
            temp_path,
            process_path,
            unprocessed_path
        )
        logging.info("File extraction completed")
        
        # Create archive subfolder and move file
        os.makedirs(os.path.join(archive_path, sub_output_folder), exist_ok=True)
        shutil.move(
            os.path.join(input_path, filename), 
            os.path.join(archive_path, sub_output_folder)
        )
        logging.info(f'File {filename} moved to archive: {os.path.join(archive_path, sub_output_folder)}')
        return True
        
    except OSError as e:
        logging.warning(e)
        return False
    except Exception as e:
        logging.error(f'Error processing file {filename}: {e}')
        os.makedirs(os.path.join(unprocessed_path, sub_output_folder), exist_ok=True)
        # Move file to unprocessed folder
        shutil.move(
            os.path.join(input_path, filename),
            os.path.join(unprocessed_path, sub_output_folder)
        )
        logging.info(f'ErrorFile {filename} moved to {os.path.join(unprocessed_path, sub_output_folder)}')
        return False


def load_config(config_path='devCode/config/config.ini'):
    """Load configuration from file"""
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


def setup_logging():
    """Configure logging"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(filename)s - %(message)s'
    )

if __name__ == '__main__':
    """Main entry point of the program"""
    setup_logging()
    
    logging.info('Initializing the config')
    config_paths = load_config()
    
    try:
        extract_input.check_create_folder(config_paths['root'])
        
        # Get sorted zip files
        sorted_files = get_sorted_zip_files(config_paths['input'])
        
        # Process each file
        for filename in sorted_files:
            process_zip_file(filename, config_paths)
            
    except Exception as e:
        logging.error(f"An error occurred: {e}")


