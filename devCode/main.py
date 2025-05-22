import os
# import logging
import shutil
from pathlib import Path
from datetime import datetime
import extract_input
from config.loadconfig import load_config
from logs.logs_handler import get_logger
from common_func.folder_operations import check_create_folder,delete_all_files,get_sorted_zip_files

logger = get_logger(__name__)

#Do not alter - this is main processing method
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
        logger.info(f'** Started Processing file: {filename} **')
        delete_all_files([temp_path, process_path]) 

        logger.info("File extraction started from input path")
        extract_input.Input_Extract(
            os.path.join(input_path, filename),
            output_path,
            error_path,
            temp_path,
            process_path,
            unprocessed_path
        )
        logger.info("File extraction completed")
        
        # Create archive subfolder and move file
        os.makedirs(os.path.join(archive_path, sub_output_folder), exist_ok=True)
        shutil.move(
            os.path.join(input_path, filename), 
            os.path.join(archive_path, sub_output_folder)
        )
        logger.info(f'File {filename} moved to archive: {os.path.join(archive_path, sub_output_folder)}')
        return True
        
    except OSError as e:
        logger.warning(e)
        return False
    except Exception as e:
        logger.error(f'Error processing file {filename}: {e}')
        os.makedirs(os.path.join(unprocessed_path, sub_output_folder), exist_ok=True)
        # Move file to unprocessed folder
        shutil.move(
            os.path.join(input_path, filename),
            os.path.join(unprocessed_path, sub_output_folder)
        )
        logger.info(f'ErrorFile {filename} moved to {os.path.join(unprocessed_path, sub_output_folder)}')
        return False

# def setup_logging():
#     """Configure logging"""
#     logging.basicConfig(
#         level=logging.INFO,
#         format='%(asctime)s - %(levelname)s - %(filename)s - %(message)s'
#     )

if __name__ == '__main__':
    """Main entry point of the program"""
    # setup_logging()
    
    logger.info('Initializing the config')
    config_paths = load_config()
    
    try:
        check_create_folder(config_paths['root'])  #change to folder_operations
        
        # Get sorted zip files
        sorted_files = get_sorted_zip_files(config_paths['input']) # change to folder_operations
        
        # Process each file
        for filename in sorted_files:
            process_zip_file(filename, config_paths) # no change
            
    except Exception as e:
        logger.error(f"An error occurred: {e}")