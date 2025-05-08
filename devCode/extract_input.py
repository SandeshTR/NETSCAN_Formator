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
from common_func.folder_operations import delete_files_in_folder,extract_zip
from core_components.jurisdictions.co  import co_region_main
from converter_modules.abbyy_integration import abby_pdf_to_docx

logger = get_logger(__name__)

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