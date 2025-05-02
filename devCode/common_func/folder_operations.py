import os
import shutil
import logging

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