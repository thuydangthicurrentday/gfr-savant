"""
Models for GOFILEROOM Downloader
Contains Document and Client classes to manage data using OOP
"""

import os
import re
import logging
import traceback
from datetime import datetime
from document_mapping import get_document_category
from utils import get_download_dir_from_env

logger = logging.getLogger(__name__)

# Global variable for base download directory (loaded from .env)
BASE_DOWNLOAD_DIR = get_download_dir_from_env()


# Custom Exceptions
class ModelError(Exception):
    """Base exception cho model errors"""
    pass


class FolderCreationError(ModelError):
    """Raised when there is an error creating folder"""
    pass


class Document:
    """Class to manage document information"""
    BE_DOWNLOAD_YEAR = 2018
    REDOWNLOAD_IF_EXISTS = True
    
    def __init__(self, document_id, file_section="", document_type="", description="", 
                 year="", document_date="", file_size="", file_type="", 
                 client_name="", client_object=None):
        """
        Initialize Document object
        
        Args:
            document_id (str): Document ID (unique)
            file_section (str): File Section
            document_type (str): Document Type
            description (str): Description
            year (str): Year
            document_date (str): Document Date
            file_size (str): File Size
            file_type (str): File Type (pdf, doc, etc.)
            client_name (str): Client Name (for creating document name)
            client_object (Client): Reference to Client object
        """
        # Thông tin cơ bản
        self.document_id = document_id
        self.file_section = file_section
        self.document_type = document_type
        self.description = description
        self.year = year
        self.document_date = document_date
        self.file_size = file_size
        self.file_type = file_type
        self.client_name = client_name
        self.client_object = client_object
        
        # Calculate category from file_section, document_type, description
        self.category_name = self.get_document_category_name()
        
        # Download status (default: "")
        self.download_status = ""  # "Success", "Error", "Warning", etc.
        self.download_description = ""
        self.download_time = ""
        
        # Calculate document names
        self.document_name_without_id = self._generate_document_name_without_id()
        self.document_name_with_id = self._generate_document_name_with_id()
        
        # Calculate attributes dependent on client_object (may be None at initialization)
        self.document_folder_path = self.get_document_folder_path()
        if self.document_folder_path:
            self.document_file_path = os.path.join(self.document_folder_path, self.document_name_with_id)
        else:
            self.document_file_path = ""
        
        self.downloadable = self.is_downloadable()
        self.file_exists = self.check_file_exists()

    def be_executed(self):
        """
        Check if document can be downloaded
        
        Returns:
            bool: True if document can be downloaded (downloadable and (file does not exist or redownload allowed))
        """
        if not self.downloadable:
            return False
        # If file does not exist, can download
        if not self.file_exists:
            return True
        # If file already exists, only download if REDOWNLOAD_IF_EXISTS = True
        return self.REDOWNLOAD_IF_EXISTS
    
    def is_downloadable(self):
        """
        Check if document can be downloaded
        
        Returns:
            bool: True if document can be downloaded (year >= BE_DOWNLOAD_YEAR or no year), False otherwise
        """
        if not self.year:
            return True
        try:
            year_number = int(self.year)
            return year_number >= self.BE_DOWNLOAD_YEAR
        except ValueError:
            return False
    
    def _generate_document_name_without_id(self):
        """
        Generate document name without _id at the end
        Format: ClientName_Year_DocumentType_Description.ext
        """
        name_items = []
        if self.client_name:
            name_items.append(str(self.client_name))
        if self.year:
            name_items.append(str(self.year))
        if self.document_type:
            name_items.append(str(self.document_type))
        if self.description:
            name_items.append(str(self.description))
        
        name = "_".join(name_items)
        # Remove special characters
        name = re.sub(r'[\\/:*?"<>|]', '', name)
        file_type = self.file_type or "pdf"
        return f"{name}.{file_type}"
    
    def _generate_document_name_with_id(self):
        """
        Generate document name with _id at the end
        Format: ClientName_Year_DocumentType_Description_document_id.ext
        """
        base_name, ext = os.path.splitext(self.document_name_without_id)
        return f"{base_name}_{self.document_id}{ext}"
    
    def get_document_category_name(self):
        """
        Get category name from file_section, document_type, description
        Uses get_document_category method from document_mapping
        
        Returns:
            str: Category name
        """
        return get_document_category(self.file_section, self.document_type, self.description)
    
    def set_download_status(self, status, description="", download_time=""):
        """
        Set download status and description
        
        Args:
            status (str): Download status ("Success", "Error", "Warning", etc.)
            description (str): Download description
            download_time (str): Download time (format: "YYYY-MM-DD HH:MM:SS")
        """
        self.download_status = status
        self.download_description = description
        if download_time:
            self.download_time = download_time
        else:
            self.download_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    def set_download_description(self, description):
        """
        Set download description
        
        Args:
            description (str): Download description
        """
        self.download_description = description
    
    def check_file_exists(self):
        """
        Check if document file already exists (check at document_file_path)
        
        Returns:
            bool: True if file exists, False otherwise
        """
        if not self.document_file_path:
            return False
        return os.path.exists(self.document_file_path) and os.path.isfile(self.document_file_path)
    
    def get_document_folder_path(self):
        """
        Return document folder path
        
        Returns:
            str: Document folder path
                - If self.year is empty: client_folder_path/category_name
                - If self.year is not empty: client_folder_path/category_name/year
        """
        if not self.client_object:
            return ""
        
        # Get client folder path
        client_folder_path = self.client_object.client_folder_path
        if not client_folder_path:
            return ""
        
        # Get category name
        category_name = self.category_name or self.get_document_category_name()
        if not category_name:
            return ""
        
        # Create path
        if self.year and self.year.strip():
            # Has year: client_folder_path/category_name/year
            return os.path.join(client_folder_path, category_name, self.year.strip())
        else:
            # No year: client_folder_path/category_name
            return os.path.join(client_folder_path, category_name)
    
    def update_paths(self):
        """
        Update paths when client_object is set or changed
        Should call this method after document is added to client
        """
        self.document_folder_path = self.get_document_folder_path()
        if self.document_folder_path:
            self.document_file_path = os.path.join(self.document_folder_path, self.document_name_with_id)
        else:
            self.document_file_path = ""
        # Update file_exists
        self.file_exists = self.check_file_exists()


class Client:
    """Class to manage client information and documents"""
    
    def __init__(self, client_name, client_number):
        """
        Initialize Client object
        
        Args:
            client_name (str): Client Name
            client_number (str): Client Number
        """
        self.client_name = client_name
        self.client_number = client_number
        
        # Create client folder name (format: client_name-client_number, remove special characters)
        self.client_folder_name = self._sanitize_folder_name(
            f"{client_number} - {client_name}"
        )
        
        # Client folder path (use global variable BASE_DOWNLOAD_DIR)
        self.client_folder_path = ""
        
        # Document list
        self.document_list = []
        
        # Download status
        self.download_client_status = ""
        self.download_client_description = ""
        self.max_total_documents = 0  # Maximum number of documents that can be downloaded
        
        # CSV download file path
        self.csv_download_file_path = ""
    
    def _sanitize_folder_name(self, folder_name):
        """
        Validate and format folder name to ensure safety for file system
        
        Args:
            folder_name (str): Original folder name
            
        Returns:
            str: Sanitized folder name
        """
        if not folder_name:
            return ""
        
        sanitized = (str(folder_name)
                    .replace("/", "")
                    .replace("\\", "")
                    .replace(":", "")
                    .replace("*", "")
                    .replace('?', "")
                    .replace('"', "")
                    .replace('<', "")
                    .replace('>', "")
                    .replace('|', ""))
        
        return sanitized
    
    def add_document(self, document):
        """
        Add document to document_list
        Check if document with same document_id already exists to avoid duplicates
        
        Args:
            document (Document): Document object
        """
        # Check if document with same document_id already exists
        existing_doc = None
        for doc in self.document_list:
            if doc.document_id == document.document_id:
                existing_doc = doc
                break
        
        # If not exists, add new
        if existing_doc is None:
            self.document_list.append(document)
            # Set client_object reference for document
            document.client_object = self
            # Update document paths
            document.update_paths()
    
    def get_number_of_downloaded_documents(self):
        """
        Count number of successfully downloaded documents
        
        Returns:
            int: Number of documents with download_status = "Success"
        """
        return sum(1 for doc in self.document_list if doc.download_status == "Success")
    
    def set_client_status(self, status):
        """
        Set client download status
        
        Args:
            status (str): Client status ("Success", "Error", "Warning", "Pending", etc.)
        """
        self.download_client_status = status
    
    def set_client_download_description(self, description):
        """
        Set client download description
        
        Args:
            description (str): Client download description
        """
        self.download_client_description = description
    
    def check_client_folder_exists(self):
        """
        Check if client folder already exists
        
        Returns:
            bool: True if folder exists, False otherwise
        """
        if not self.client_folder_path:
            return False
        return os.path.exists(self.client_folder_path) and os.path.isdir(self.client_folder_path)
    
    def check_category_folder_exists(self, category):
        """
        Check if category folder already exists in client folder
        
        Args:
            category (str): Category name
            
        Returns:
            bool: True if folder exists, False otherwise
        """
        if not self.client_folder_path:
            return False
        category_path = os.path.join(self.client_folder_path, category)
        return os.path.exists(category_path) and os.path.isdir(category_path)
    
    def check_year_folder_exists(self, category, year):
        """
        Check if year folder already exists in category folder
        
        Args:
            category (str): Category name
            year (str): Year
            
        Returns:
            bool: True if folder exists, False otherwise
        """
        if not self.client_folder_path or not year:
            return False
        year_path = os.path.join(self.client_folder_path, category, year)
        return os.path.exists(year_path) and os.path.isdir(year_path)
    
    def create_client_folder(self):
        """
        Create client folder if it doesn't exist
        
        Returns:
            str: Full path to created folder
            
        Raises:
            FolderCreationError: If there is an error creating client folder
        """
        base_download_dir = BASE_DOWNLOAD_DIR
        if not base_download_dir:
            error_msg = f"BASE_DOWNLOAD_DIR is not configured"
            logger.error(error_msg)
            raise FolderCreationError(error_msg)
        
        if not self.client_folder_name:
            error_msg = f"Client folder name is invalid"
            logger.error(error_msg)
            raise FolderCreationError(error_msg)
        
        try:
            full_path = os.path.join(base_download_dir, self.client_folder_name)
            
            if not os.path.exists(full_path):
                os.makedirs(full_path, exist_ok=True)
            
            self.client_folder_path = full_path
            logger.debug(f"Created client folder: {full_path}")
            return full_path
            
        except Exception as e:
            error_msg = f"Error creating client folder: {str(e)}"
            logger.error(error_msg)
            logger.error(traceback.format_exc())
            raise FolderCreationError(error_msg) from e
    
    def create_category_folders(self):
        """
        Create all category folders in client folder if they don't exist
        
        Raises:
            FolderCreationError: If there is an error creating category folders
        """
        if not self.client_folder_path or not os.path.exists(self.client_folder_path):
            error_msg = f"Client folder does not exist: {self.client_folder_path}"
            logger.error(error_msg)
            raise FolderCreationError(error_msg)
        
        try:
            from document_mapping import get_all_categories
            categories = get_all_categories()
            
            for category in categories:
                category_path = os.path.join(self.client_folder_path, category)
                if not os.path.exists(category_path):
                    os.makedirs(category_path, exist_ok=True)
            
            logger.debug(f"Created category folders for client: {self.client_name}")
            
        except Exception as e:
            error_msg = f"Error creating category folders: {str(e)}"
            logger.error(error_msg)
            logger.error(traceback.format_exc())
            raise FolderCreationError(error_msg) from e
    
    def initialize_folders(self):
        """
        Initialize client folder and category folders
        Should be called after client is successfully found/verified
        
        Only sets self.client_folder_path after BOTH steps succeed.
        If any step fails, self.client_folder_path remains empty.
        
        Raises:
            FolderCreationError: If there is an error creating folders
        """
        try:
            # Create client folder (this will temporarily set self.client_folder_path internally)
            self.create_client_folder()
            # At this point, self.client_folder_path is set by create_client_folder()
            
            # Create category folders (requires client_folder_path to be set)
            self.create_category_folders()
            
            # Only log success if both steps completed
            # self.client_folder_path is already set by create_client_folder() if successful
            logger.info(f"Initialized folders for client: {self.client_name} ({self.client_number})")
            
        except FolderCreationError as e:
            # If any step fails, reset to empty string
            # This ensures self.client_folder_path is only set when BOTH steps succeed
            self.client_folder_path = ""
            error_msg = f"Cannot create folders for client {self.client_name} ({self.client_number}): {str(e)}"
            logger.error(error_msg)
            raise FolderCreationError(error_msg) from e
    
    def create_year_folder(self, category, year):
        """
        Create year folder in category folder if it doesn't exist
        
        Args:
            category (str): Category name
            year (str): Year
            
        Returns:
            str: Full path to created year folder
            
        Raises:
            FolderCreationError: If there is an error creating year folder
        """
        if not self.client_folder_path or not year:
            error_msg = f"Client folder path or year is invalid: folder={self.client_folder_path}, year={year}"
            logger.error(error_msg)
            raise FolderCreationError(error_msg)
        
        if not os.path.exists(self.client_folder_path):
            error_msg = f"Client folder does not exist: {self.client_folder_path}"
            logger.error(error_msg)
            raise FolderCreationError(error_msg)
        
        try:
            category_path = os.path.join(self.client_folder_path, category)
            if not os.path.exists(category_path):
                os.makedirs(category_path, exist_ok=True)
            
            year_path = os.path.join(category_path, year)
            if not os.path.exists(year_path):
                os.makedirs(year_path, exist_ok=True)
            
            logger.debug(f"Created year folder: {year_path}")
            return year_path
            
        except Exception as e:
            error_msg = f"Error creating year folder: {str(e)}"
            logger.error(error_msg)
            logger.error(traceback.format_exc())
            raise FolderCreationError(error_msg) from e

