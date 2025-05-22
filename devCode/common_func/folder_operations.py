import shutil
import logging
from logs.logs_handler import get_logger
import tkinter as tk
from tkinter import messagebox
import os
import zipfile
from pathlib import Path
from datetime import datetime

logger = get_logger(__name__)

def delete_folder(folder_path):
    """
    Delete a folder and all its contents (subfolders and files)
    
    """
    try:
        if not os.path.exists(folder_path):
            print(f"The folder '{folder_path}' does not exist.")
            return False
            
        if os.path.isdir(folder_path):
            # Remove the entire directory tree
            shutil.rmtree(folder_path)
            print(f"Successfully deleted '{folder_path}' and all its contents.")
            return True
        else:
            print(f"'{folder_path}' is not a directory.")
            return False
    except Exception as e:
        print(f"Error while deleting '{folder_path}': {e}")
        return False


def delete_files_in_folder(folder_path):
    """
    Delete all files in a folder without deleting the folder itself
    """
    try:
        if not os.path.exists(folder_path):
            print(f"The folder '{folder_path}' does not exist.")
            return False
            
        if not os.path.isdir(folder_path):
            print(f"'{folder_path}' is not a directory.")
            return False
            
        file_count = 0
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            if os.path.isfile(file_path):
                os.unlink(file_path)
                file_count += 1
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
                file_count += 1
                
        print(f"Successfully deleted {file_count} items from '{folder_path}'.")
        return True
    except Exception as e:
        print(f"Error while deleting files in '{folder_path}': {e}")
        return False
    
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


def check_create_folder(base_folder_path):
    """
    Check if base folder exists, create it if not.
    Also creates standard subfolders and notifies user if input folder was created.
    """
    if not os.path.exists(base_folder_path):
        os.mkdir(base_folder_path)
        logger.info(f'Created base folder: {base_folder_path}')

    subfolders = ['Input', 'Output', 'Process', 'Temp', 'Unprocessed']
    input_folder_created = False
    input_folder_path = ""

    for folder in subfolders:
        folder_path = os.path.join(base_folder_path, folder)
        if os.path.exists(folder_path):
            logger.info(f'Subfolder already present: {folder_path}')
        else:
            try:
                os.makedirs(folder_path)
                logger.info(f'Created folder: {folder_path}')
                if folder == 'Input':
                    input_folder_created = True
                    input_folder_path = folder_path
            except OSError as e:
                logger.error(f'Error creating subfolder {folder_path}: {e}')

    if input_folder_created:
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        messagebox.showinfo("Input Folder Created", 
                            f"Input folder was not present and has been created.\n"
                            f"Input folder path: {input_folder_path}\n"
                            f"Please place zip folders in the Input folder.")
        root.destroy()

def extract_zip(zip_path, extract_to, unprocessed_path):
    """Extract ZIP file and handle nested ZIP files."""
    logger.info(f"Extracting ZIP file: {zip_path} to {extract_to}")
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_to)
            logger.info(f"Extracted {zip_path} to {extract_to}")

            for root, dirs, files in os.walk(extract_to):
                for file_name in files:
                    if file_name.lower().endswith('.zip'):
                        file_path = os.path.join(root, file_name)
                        if zipfile.is_zipfile(file_path):
                            logger.info(f"Found nested ZIP file: {file_path}")
                            nested_extract_to = os.path.join(root, os.path.splitext(file_name)[0])
                            os.makedirs(nested_extract_to, exist_ok=True)
                            extract_zip(file_path, nested_extract_to, unprocessed_path)
                            os.remove(file_path)

        logger.info(f"Completed extraction of {zip_path}")
    except zipfile.BadZipFile:
        logger.error(f"Bad ZIP file: {zip_path}")    
        bad_zip_dir = os.path.join(unprocessed_path, datetime.now().strftime("%B_%d_%Y").upper())
        os.makedirs(bad_zip_dir, exist_ok=True)
        shutil.move(zip_path, os.path.join(bad_zip_dir, os.path.basename(zip_path)))
        logger.info(f"Moved bad ZIP file to: {bad_zip_dir}")
    except Exception as e:
        logger.error(f"An error occurred while extracting {zip_path}: {str(e)}")

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
        logger.warning("No .zip files present in input folder")
        
    return sorted_files

def clean_up_temp_file(temp_file_path):
    """
    Remove the temporary DOCX file if it exists.
    """
    if os.path.exists(temp_file_path):
        os.remove(temp_file_path)
        
# def delete_all_files(folder_path):
#     """Delete all files in the specified folder."""
#     for filename in os.listdir(folder_path):
#         file_path = os.path.join(folder_path, filename)
#         # Check if it is a file before deleting
#         if os.path.isfile(file_path):
#             try:
#                 os.remove(file_path)
#                 logging.info(f"Deleted file: {file_path}")
#             except Exception as e:
#                 logging.error(f"Unable to delete file {file_path}: {e}")
#                 try:
#                     os.chmod(file_path, 0o777)  # Change the file permission to writable
#                     os.remove(file_path)
#                     logging.info(f"Forcefully deleted file: {file_path}")
#                 except Exception as e:
#                     logging.error(f"Failed to forcefully delete file {file_path}: {e}")


# def delete_all_files_and_folders(folder_path):
#     """Delete all files and folders within the specified folder."""
#     for root, dirs, files in os.walk(folder_path, topdown=False):
#         # First, remove all files in the current directory
#         for file in files:
#             file_path = os.path.join(root, file)
#             try:
#                 os.remove(file_path)
#                 logging.info(f"Deleted file: {file_path}")
#             except Exception as e:
#                 logging.error(f"Failed to delete file {file_path}: {e}")
        
#         # Then, remove all subdirectories
#         for dir in dirs:
#             dir_path = os.path.join(root, dir)
#             try:
#                 shutil.rmtree(dir_path)
#                 logging.info(f"Deleted directory: {dir_path}")
#             except Exception as e:
#                 logging.error(f"Failed to delete directory {dir_path}: {e}")
