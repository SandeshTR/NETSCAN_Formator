import zipfile
import os
import glob
import shutil
import time
import gc
import tkinter as tk
from tkinter import messagebox
from datetime import datetime
from logs.logs_handler import get_logger

# Local imports
import core_components.jurisdictions.ca as HTMLtoWord
from converter_modules.com_integration.com_word_format_converter import convert_file_to_docx
from core_components.generic_instruction.generic_instructions import format_document
from common_func.folder_operations import delete_files_in_folder
from core_components.jurisdictions.co  import co_region_main
from converter_modules.abbyy_integration import abby_pdf_to_docx

logger = get_logger(__name__)

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


def extract_next_two_chars(filename, keyword):
    """Extract the next 2 characters after the keyword from the given filename."""
    start_index = filename.find(keyword)
    if start_index == -1:
        logger.warning(f"Keyword '{keyword}' not found in filename: {filename}")
        return None
    start_index += len(keyword)
    next_two_chars = filename[start_index:start_index + 2]
    logger.debug(f"Extracted characters: {next_two_chars}")
    return next_two_chars


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


def process_file(file_path, jurisdiction, output_path, error_path, temp_path, process_path):
    """Process a single file based on its extension and jurisdiction."""
    delete_files_in_folder(process_path)
    region_code = jurisdiction.lower() 
    
    logger.info(f"Processing file: {file_path}")
    
    # Create output directory path
    output_dir = os.path.join("C:\\File\\NETSCAN\\Output", os.path.basename(os.path.dirname(file_path)))
    os.makedirs(output_dir, exist_ok=True)

    # Get base filename without extension
    base_filename = os.path.splitext(os.path.basename(file_path))[0]
    process_docx_path = os.path.join(process_path, f"{base_filename}.docx")

    # Convert files to docx format based on file extension
    if file_path.lower().endswith('.pdf'):
        abby_pdf_to_docx.Run(file_path=file_path, output_path=process_docx_path)
        logger.info(f'Converted PDF to DOCX: {process_docx_path}')

    elif file_path.lower().endswith('.doc'):
        convert_file_to_docx(file_path=file_path, output_path=process_docx_path)
        logger.info(f'Converted DOC to DOCX: {process_docx_path}')

    elif file_path.lower().endswith('.docx'):
        shutil.copy2(file_path, process_docx_path)
        logger.info(f'Copied DOCX file to process directory: {process_docx_path}')
    
    # Process based on jurisdiction
    if jurisdiction.lower() == "ca":
        pass  # CA processing code commented out
    elif jurisdiction.lower() == "co":
        if file_path.lower().endswith('.html'):
            HTMLtoWord.HTMLtoWord(file_path, process_docx_path)
        elif os.path.splitext(file_path)[1].strip().lower() in ['.pdf', '.doc', '.docx']:
            co_region_main.main_co_files(
                input_docx_path=process_docx_path,
                source_file_path=file_path
            )
    
    # Apply generic formatting rules
    format_document(doc_path=process_docx_path, region_code=region_code)
    
    # Move processed file to output folder (as .doc)
    output_file_path = os.path.join(output_dir, f"{base_filename}.doc")
    shutil.move(process_docx_path, output_file_path)
    logger.info(f"Moved processed file to: {output_file_path}")

    # Clean up original file
    if os.path.exists(file_path):
        os.remove(file_path)
        logger.info(f"Original file deleted: {file_path}")

    # Clean up any remaining process files
    if os.path.exists(process_docx_path):
        os.remove(process_docx_path)
        logger.info('Deleted processed file present in process folder')

    gc.collect()
    time.sleep(3)


def loop_through_folders(base_directory, jurisdiction, output_path, error_path, temp_path, process_path):
    """Loop through folders and process PDF, DOC, and ZIP files."""
    logger.info(f"Looping through folders in base directory: {base_directory}")
    
    if not os.listdir(base_directory):
        os.rmdir(base_directory)
        raise OSError('No files present in folder')

    file_patterns = ["*.pdf", "*.doc", "*.html", "*.docx"]
    
    # Recursively loop through directories
    for root, dirs, files in os.walk(base_directory):
        for pattern in file_patterns:
            # Generate file paths matching the pattern
            for file_path in glob.glob(os.path.join(root, pattern)):
                logger.info(f'Found file: {file_path}')

                try:
                    process_file(file_path, jurisdiction, output_path, error_path, temp_path, process_path)
                except Exception as e:
                    logger.error(f'Error processing file {file_path}: {e}')
                    # Move file to exception directory
                    output_dir = os.path.join("C:\\File\\NETSCAN\\Output", os.path.basename(os.path.dirname(file_path)))
                    exception_dir = os.path.join(output_dir, "Exception")
                    os.makedirs(exception_dir, exist_ok=True)
                    
                    exception_file = os.path.join(exception_dir, os.path.basename(file_path))
                    shutil.move(file_path, exception_file)
                    logger.info(f'Moved file to exception directory: {exception_file}')


def Input_Extract(input_path, output_path, error_path, temp_path, process_path, unprocessed_path):
    """Extract and process files from input ZIP file."""
    jurisdiction = ""

    if input_path.lower().endswith('.zip'):
        jurisdiction = extract_next_two_chars(input_path, 'NETSCAN_')

        # Directory to extract files to, named after the root ZIP file
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        output_dir = os.path.join(os.path.dirname(input_path), base_name)

        # Ensure the output directory exists
        os.makedirs(output_dir, exist_ok=True)

        # Start extraction process
        extract_zip(input_path, output_dir, unprocessed_path)
        loop_through_folders(output_dir, jurisdiction, output_path, error_path, temp_path, process_path)
        shutil.rmtree(output_dir)
    else:
        logger.warning(f"Input path does not end with .zip: {input_path}")