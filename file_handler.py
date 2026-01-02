"""
File Handler for GOFILEROOM Downloader
Handles all file operations after downloading from browser
"""

import os
import shutil
import zipfile
import logging
import time
from datetime import datetime

logger = logging.getLogger(__name__)


# Custom Exceptions
class FileHandlerError(Exception):
    """Base exception cho file handler errors"""
    pass


class FileNotFoundError(FileHandlerError):
    """Raised when file does not exist"""
    pass


class FileOperationError(FileHandlerError):
    """Raised when there is an error in file operations"""
    pass


class ZipFileError(FileHandlerError):
    """Raised when there is an error with ZIP file"""
    pass


class FileDownloadTimeoutError(FileHandlerError):
    """Raised when timeout waiting for file download"""
    pass


def rename_file_with_doc_id(file_path, doc_id):
    """
    Rename file to add document_id at the end of the name
    
    Args:
        file_path (str): Original file path
        doc_id (str): Document ID to add to file name
        
    Returns:
        str: New file path after renaming
        
    Raises:
        FileNotFoundError: If file does not exist
        FileOperationError: If there is an error renaming file
    """
    if not os.path.exists(file_path):
        error_msg = f"File does not exist: {file_path}"
        raise FileNotFoundError(error_msg)
    
    try:
        # Split file name and extension
        base_name, ext = os.path.splitext(os.path.basename(file_path))
        new_file_name = f"{base_name}_{doc_id}{ext}"
        
        # New file path (same directory as original file)
        file_dir = os.path.dirname(file_path)
        new_file_path = os.path.join(file_dir, new_file_name)
        
        # Rename file
        os.rename(file_path, new_file_path)
        logger.info(f"Renamed file: {os.path.basename(file_path)} -> {new_file_name}")
        
        return new_file_path
        
    except FileNotFoundError:
        raise
    except Exception as e:
        error_msg = f"Error renaming file: {str(e)}"
        raise FileOperationError(error_msg) from e


def move_file(source_path, destination_path):
    """
    Move file from source_path to destination_path
    
    Args:
        source_path (str): Source file path
        destination_path (str): Destination file path
        
    Raises:
        FileNotFoundError: If source file does not exist
        FileOperationError: If there is an error moving file
    """
    if not os.path.exists(source_path):
        error_msg = f"Source file does not exist: {source_path}"
        raise FileNotFoundError(error_msg)
    
    try:
        # Create destination directory if it doesn't exist
        dest_dir = os.path.dirname(destination_path)
        if dest_dir and not os.path.exists(dest_dir):
            os.makedirs(dest_dir, exist_ok=True)
        
        # Move file
        shutil.move(source_path, destination_path)
        logger.info(f"Moved file: {os.path.basename(source_path)} -> {destination_path}")
        
    except FileNotFoundError:
        raise
    except Exception as e:
        error_msg = f"Error moving file: {str(e)}"
        raise FileOperationError(error_msg) from e


def move_csv_to_storage(csv_file_path, csv_dir):
    """
    Move CSV file to 0_csv_ directory
    
    Args:
        csv_file_path (str): CSV file path in download_dir
        csv_dir (str): Directory to store CSV files (0_csv_)
        
    Returns:
        str: CSV file path after moving
        
    Raises:
        FileNotFoundError: If CSV file does not exist
        FileOperationError: If there is an error moving CSV file
    """
    csv_filename = os.path.basename(csv_file_path)
    csv_dest_path = os.path.join(csv_dir, csv_filename)
    
    # Nếu file đã tồn tại, đổi tên với timestamp
    if os.path.exists(csv_dest_path):
        base, ext = os.path.splitext(csv_filename)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        csv_filename = f"{base}_{timestamp}{ext}"
        csv_dest_path = os.path.join(csv_dir, csv_filename)
    
    # Sử dụng move_file để thực hiện move
    move_file(csv_file_path, csv_dest_path)
    return csv_dest_path


def rename_csv_file(csv_file_path, client_folder_name):
    """
    Rename CSV file to format: Search_<client_folder_name>.csv
    
    Args:
        csv_file_path (str): Current CSV file path
        client_folder_name (str): Client folder name to use in new file name
        
    Returns:
        str: New CSV file path after renaming
        
    Raises:
        FileNotFoundError: If CSV file does not exist
        FileOperationError: If there is an error renaming CSV file
    """
    if not os.path.exists(csv_file_path):
        error_msg = f"CSV file does not exist: {csv_file_path}"
        raise FileNotFoundError(error_msg)
    
    try:
        # Get directory and extension
        file_dir = os.path.dirname(csv_file_path)
        _, ext = os.path.splitext(csv_file_path)
        
        # Create new file name: Search_<client_folder_name>.csv
        new_file_name = f"Search_{client_folder_name}{ext}"
        new_file_path = os.path.join(file_dir, new_file_name)
        
        # If file with new name already exists, add timestamp
        if os.path.exists(new_file_path):
            base_name = f"Search_{client_folder_name}"
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            new_file_name = f"{base_name}_{timestamp}{ext}"
            new_file_path = os.path.join(file_dir, new_file_name)
        
        # Rename file
        os.rename(csv_file_path, new_file_path)
        logger.info(f"Renamed CSV file: {os.path.basename(csv_file_path)} -> {new_file_name}")
        
        return new_file_path
        
    except FileNotFoundError:
        raise
    except Exception as e:
        error_msg = f"Error renaming CSV file: {str(e)}"
        raise FileOperationError(error_msg) from e


def move_zip_to_storage(zip_file_path, zip_client_folder_path):
    """
    Move ZIP file to zip_client_folder_path directory
    
    Args:
        zip_file_path (str): ZIP file path in download_dir
        zip_client_folder_path (str): Directory path to store ZIP for client
        
    Returns:
        str: ZIP file path after moving
        
    Raises:
        FileNotFoundError: If ZIP file does not exist
        FileOperationError: If there is an error moving ZIP file
    """
    zip_filename = os.path.basename(zip_file_path)
    zip_dest_path = os.path.join(zip_client_folder_path, zip_filename)
    
    # Sử dụng move_file để thực hiện move
    move_file(zip_file_path, zip_dest_path)
    return zip_dest_path


def extract_zip(zip_path, extract_dir):
    """
    Extract ZIP file to extract_dir
    
    Args:
        zip_path (str): ZIP file path
        extract_dir (str): Extraction directory
        
    Raises:
        FileNotFoundError: If ZIP file does not exist
        ZipFileError: If ZIP file is corrupted or there is an error extracting
    """
    if not os.path.exists(zip_path):
        error_msg = f"ZIP file does not exist: {zip_path}"
        raise FileNotFoundError(error_msg)
    
    try:
        # Create extraction directory if it doesn't exist
        if not os.path.exists(extract_dir):
            os.makedirs(extract_dir, exist_ok=True)
        
        # Extract
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)
        
        logger.info(f"Extracted ZIP: {zip_path} to {extract_dir}")
        
    except zipfile.BadZipFile as e:
        error_msg = f"ZIP file is corrupted: {zip_path}"
        raise ZipFileError(error_msg) from e
    except Exception as e:
        error_msg = f"Error extracting ZIP: {str(e)}"
        raise ZipFileError(error_msg) from e


def clean_download_dir(download_dir):
    """
    Delete all files in download_dir but keep subdirectories
    
    Args:
        download_dir (str): Download directory path
        
    Raises:
        FileOperationError: If there is an error deleting files
    """
    if not os.path.exists(download_dir):
        error_msg = f"Download directory does not exist: {download_dir}"
        raise FileOperationError(error_msg)
    
    try:
        deleted_count = 0
        for item in os.listdir(download_dir):
            item_path = os.path.join(download_dir, item)
            
            # Only delete files, not directories
            if os.path.isfile(item_path):
                try:
                    os.remove(item_path)
                    deleted_count += 1
                    logger.debug(f"Deleted file: {item}")
                except Exception as e:
                    logger.warning(f"Cannot delete file {item}: {str(e)}")
        
        if deleted_count > 0:
            logger.info(f"Deleted {deleted_count} file(s) in download_dir: {download_dir}")
        else:
            logger.debug(f"No files to delete in download_dir: {download_dir}")
            
    except Exception as e:
        error_msg = f"Error deleting files in download_dir: {str(e)}"
        raise FileOperationError(error_msg) from e


def remove_file(file_path):
    """
    Remove file
    
    Args:
        file_path (str): File path to remove
        
    Raises:
        FileOperationError: If there is an error removing file (does not raise if file does not exist)
    """
    try:
        if os.path.exists(file_path):
            os.remove(file_path)
            logger.debug(f"Removed file: {file_path}")
    except Exception as e:
        error_msg = f"Error removing file: {str(e)}"
        logger.warning(error_msg)
        raise FileOperationError(error_msg) from e


def find_file_in_zip_folder(zip_folder_path, doc_id, expected_name_pattern):
    """
    Find file in zip folder based on doc_id and expected_name_pattern
    
    Args:
        zip_folder_path (str): Path to extracted zip folder
        doc_id (str): Document ID
        expected_name_pattern (str): Expected file name pattern (may contain doc_id)
        
    Returns:
        str: Path to found file
        
    Raises:
        FileNotFoundError: If zip folder does not exist or file not found
        FileOperationError: If there is an error finding file
    """
    if not os.path.exists(zip_folder_path):
        error_msg = f"Zip folder does not exist: {zip_folder_path}"
        raise FileNotFoundError(error_msg)
    
    try:
        # Find file in zip folder
        for root, dirs, files in os.walk(zip_folder_path):
            for file in files:
                # File from zip has format: expected_name_document_id.ext
                if doc_id in file and expected_name_pattern in file:
                    found_file = os.path.join(root, file)
                    return found_file
        
        # File not found
        error_msg = f"File not found with doc_id={doc_id} and pattern={expected_name_pattern} in {zip_folder_path}"
        raise FileNotFoundError(error_msg)
        
    except FileNotFoundError:
        raise
    except Exception as e:
        error_msg = f"Error finding file in zip folder: {str(e)}"
        raise FileOperationError(error_msg) from e


def wait_for_file_download(folder_path, expected_extension="", timeout=120):
    """
    Wait for file to be downloaded in folder_path
    
    Args:
        folder_path (str): Download directory
        expected_extension (str): Expected extension (e.g., ".csv", ".zip")
        timeout (int): Maximum wait time (seconds)
        
    Returns:
        str: Path to downloaded file
        
    Raises:
        FileDownloadTimeoutError: If timeout and file not found
        FileOperationError: If there is an error waiting for file download
    """
    try:
        logger.info(f"Waiting for file to download (timeout: {timeout}s)...")
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            # Check for temporary files
            temp_files = [f for f in os.listdir(folder_path) 
                         if f.endswith(('.crdownload', '.tmp'))]
            if temp_files:
                time.sleep(2)
                continue
            
            # Get list of files
            if expected_extension:
                files = [f for f in os.listdir(folder_path)
                        if f.endswith(expected_extension) and 
                        os.path.isfile(os.path.join(folder_path, f))]
            else:
                files = [f for f in os.listdir(folder_path)
                        if os.path.isfile(os.path.join(folder_path, f))]
            
            if files:
                # Get latest file
                full_paths = [os.path.join(folder_path, f) for f in files]
                latest_file = max(full_paths, key=os.path.getmtime)
                
                # Check if file download is complete
                if not os.path.exists(latest_file + ".crdownload"):
                    logger.info(f"File download completed: {os.path.basename(latest_file)}")
                    return latest_file
            
            time.sleep(2)
        
        error_msg = f"Timeout: File not found after {timeout} seconds"
        raise FileDownloadTimeoutError(error_msg)
        
    except FileDownloadTimeoutError:
        raise
    except Exception as e:
        error_msg = f"Error waiting for file download: {str(e)}"
        raise FileOperationError(error_msg) from e

