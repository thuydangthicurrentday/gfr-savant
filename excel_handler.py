"""
Excel Handler for GOFILEROOM Downloader
Handles all Excel file operations
"""

import logging

logger = logging.getLogger(__name__)


# Custom Exceptions
class ExcelHandlerError(Exception):
    """Base exception cho Excel handler errors"""
    pass


class ExcelHeaderError(ExcelHandlerError):
    """Raised when header or column not found in Excel"""
    pass


class ExcelOperationError(ExcelHandlerError):
    """Raised when there is an error in Excel operations"""
    pass


class ExcelSaveError(ExcelHandlerError):
    """Raised when there is an error saving Excel file"""
    pass


class ExcelHandler:
    """Class to handle all Excel operations"""
    
    def __init__(self, workbook, client_list_sheet, document_list_sheet, excel_file_path):
        """
        Initialize ExcelHandler
        
        Args:
            workbook: openpyxl workbook object
            client_list_sheet: openpyxl worksheet for client list
            document_list_sheet: openpyxl worksheet for document list
            excel_file_path (str): Excel file path
        """
        self.workbook = workbook
        self.client_list_sheet = client_list_sheet
        self.document_list_sheet = document_list_sheet
        self.excel_file_path = excel_file_path
        
        # List attributes
        self.client_list = []
        self.document_list = []
        
        # Read lists on initialization
        self.client_list = self.get_client_list()
        self.document_list = self.get_document_list()
    
    def get_client_header_indices(self):
        """
        Read header of client_list_sheet to get column indices
        
        Returns:
            dict: Dictionary containing column indices
            
        Raises:
            ExcelHeaderError: If required columns not found in header
            ExcelOperationError: If there is an error reading header
        """
        try:
            header_row = [cell.value for cell in self.client_list_sheet[1]]
            
            indices = {
                'status': header_row.index('Status'),
                'description': header_row.index('Description'),
                'client_name': header_row.index('Client Name'),
                'client_number': header_row.index('Client Number'),
                'total_documents': header_row.index('Total Documents'),
                'num_files_downloaded': header_row.index('Number Of Files Downloaded'),
                'client_folder_path': header_row.index('Client Folder Path')
            }
            
            # Add client_email if exists
            try:
                indices['client_email'] = header_row.index('Client Email')
            except ValueError:
                indices['client_email'] = None
            
            return indices
            
        except ValueError as e:
            error_msg = f"Column not found in client_list_sheet header: {str(e)}"
            raise ExcelHeaderError(error_msg) from e
        except Exception as e:
            error_msg = f"Error reading client_list_sheet header: {str(e)}"
            raise ExcelOperationError(error_msg) from e
    
    def get_document_header_indices(self):
        """
        Read header of document_list_sheet to get column indices
        
        Returns:
            dict: Dictionary containing column indices
            
        Raises:
            ExcelHeaderError: If required columns not found in header
            ExcelOperationError: If there is an error reading header
        """
        try:
            header_row = [cell.value for cell in self.document_list_sheet[1]]
            
            indices = {
                'download_status': header_row.index('Download Status'),
                'download_desc': header_row.index('Download Description'),
                'client_name': header_row.index('Client Name'),
                'client_number': header_row.index('Client Number'),
                'file_name': header_row.index('File Name'),
                'file_path': header_row.index('File Path'),
                'folder_category': header_row.index('Folder Category'),
                'file_section': header_row.index('File Section'),
                'doc_type': header_row.index('Document Type'),
                'description': header_row.index('Description'),
                'year': header_row.index('Year'),
                'doc_date': header_row.index('Document Date'),
                'file_size': header_row.index('File Size'),
                'doc_id': header_row.index('Document ID'),
                'file_type': header_row.index('File Type'),
                'download_time': header_row.index('Download time')
            }
            
            return indices
            
        except ValueError as e:
            error_msg = f"Column not found in document_list_sheet header: {str(e)}"
            raise ExcelHeaderError(error_msg) from e
        except Exception as e:
            error_msg = f"Error reading document_list_sheet header: {str(e)}"
            raise ExcelOperationError(error_msg) from e
    
    def get_client_list(self, status_filter=None):
        """
        Read client list from client_list_sheet
        
        Args:
            status_filter (str): Filter clients by status (default: None - get all)
        
        Returns:
            list: Client list, each item is a dict containing client info and cell references
            
        Raises:
            ExcelHeaderError: If header cannot be read
            ExcelOperationError: If there is an error reading client list
        """
        try:
            indices = self.get_client_header_indices()
            
            client_list = []
            for row_idx, row in enumerate(self.client_list_sheet.iter_rows(min_row=2, values_only=False), start=2):
                try:
                    client_name = str(row[indices['client_name']].value or "").strip()
                    client_number = str(row[indices['client_number']].value or "").strip()
                    
                    if not client_name or not client_number:
                        continue
                    
                    # Filter by status if status_filter is provided
                    if status_filter is not None:
                        status_value = str(row[indices['status']].value or "").strip()
                        if status_value != status_filter:
                            continue
                    
                    client_info = {
                        'row_index': row_idx,
                        'status_cell': row[indices['status']],
                        'description_cell': row[indices['description']],
                        'client_name_cell': row[indices['client_name']],
                        'client_number_cell': row[indices['client_number']],
                        'total_documents_cell': row[indices['total_documents']],
                        'num_files_downloaded_cell': row[indices['num_files_downloaded']],
                        'client_folder_path_cell': row[indices['client_folder_path']],
                        'client_name': client_name,
                        'client_number': client_number,
                    }
                    
                    # Add client_email if exists
                    if indices.get('client_email') is not None:
                        client_info['client_email_cell'] = row[indices['client_email']]
                    
                    client_list.append(client_info)
                except Exception as e:
                    logger.warning(f"Error reading row {row_idx} in client_list_sheet: {str(e)}")
                    continue
            
            if status_filter:
                logger.info(f"Read {len(client_list)} clients with Status = {status_filter} from client_list_sheet")
            else:
                logger.info(f"Read {len(client_list)} clients from client_list_sheet")
            return client_list
            
        except (ExcelHeaderError, ExcelOperationError):
            raise
        except Exception as e:
            error_msg = f"Error reading client list: {str(e)}"
            raise ExcelOperationError(error_msg) from e
    
    def get_document_list(self):
        """
        Read document list from document_list_sheet
        
        Returns:
            list: Document list, each item is a dict containing document info
            
        Raises:
            ExcelHeaderError: If header cannot be read
            ExcelOperationError: If there is an error reading document list
        """
        try:
            indices = self.get_document_header_indices()
            
            document_list = []
            for row_idx, row in enumerate(self.document_list_sheet.iter_rows(min_row=2, values_only=False), start=2):
                try:
                    doc_id = str(row[indices['doc_id']].value or "").strip()
                    client_name = str(row[indices['client_name']].value or "").strip()
                    client_number = str(row[indices['client_number']].value or "").strip()
                    
                    if not doc_id or not client_name or not client_number:
                        continue
                    
                    doc_info = {
                        'row_index': row_idx,
                        'doc_id': doc_id,
                        'client_name': client_name,
                        'client_number': client_number,
                        'download_status': str(row[indices['download_status']].value or "").strip(),
                        'download_desc': str(row[indices['download_desc']].value or "").strip(),
                        'file_name': str(row[indices['file_name']].value or "").strip(),
                        'file_path': str(row[indices['file_path']].value or "").strip(),
                        'folder_category': str(row[indices['folder_category']].value or "").strip(),
                        'file_section': str(row[indices['file_section']].value or "").strip(),
                        'doc_type': str(row[indices['doc_type']].value or "").strip(),
                        'description': str(row[indices['description']].value or "").strip(),
                        'year': str(row[indices['year']].value or "").strip(),
                        'doc_date': str(row[indices['doc_date']].value or "").strip(),
                        'file_size': str(row[indices['file_size']].value or "").strip(),
                        'file_type': str(row[indices['file_type']].value or "").strip(),
                        'download_time': str(row[indices['download_time']].value or "").strip(),
                    }
                    
                    document_list.append(doc_info)
                except Exception as e:
                    logger.warning(f"Error reading row {row_idx} in document_list_sheet: {str(e)}")
                    continue
            
            logger.info(f"Read {len(document_list)} documents from document_list_sheet")
            return document_list
            
        except (ExcelHeaderError, ExcelOperationError):
            raise
        except Exception as e:
            error_msg = f"Error reading document list: {str(e)}"
            raise ExcelOperationError(error_msg) from e
    
    def get_client_row_index(self, client_name, client_number):
        """
        Find row_index of client in client_list_sheet
        
        Args:
            client_name (str): Client Name
            client_number (str): Client Number
            
        Returns:
            int: Row index if found, None if not found
        """
        for client_info in self.client_list:
            if (client_info['client_name'] == client_name and 
                client_info['client_number'] == client_number):
                return client_info['row_index']
        return None
    
    def get_document_row_index(self, client_name, client_number, document_id):
        """
        Find row_index of document in document_list_sheet
        
        Args:
            client_name (str): Client Name
            client_number (str): Client Number
            document_id (str): Document ID
            
        Returns:
            int: Row index if found, None if not found
        """
        for doc_info in self.document_list:
            if (doc_info['client_name'] == client_name and 
                doc_info['client_number'] == client_number and
                doc_info['doc_id'] == document_id):
                return doc_info['row_index']
        return None
    
    def update_client_row(self, row_index, status=None, description=None, 
                         total_documents=None, num_files_downloaded=None, 
                         client_folder_path=None):
        """
        Update data to row by row_index in client_list_sheet
        
        Args:
            row_index (int): Row index to update
            status (str): Status (optional)
            description (str): Description (optional)
            total_documents (int): Total documents (optional)
            num_files_downloaded (int): Number of files downloaded (optional)
            client_folder_path (str): Client folder path (optional)
            
        Raises:
            ExcelHeaderError: If header cannot be read
            ExcelOperationError: If there is an error updating client row
        """
        try:
            indices = self.get_client_header_indices()
            
            row = self.client_list_sheet[row_index]
            
            if status is not None:
                row[indices['status']].value = status
            if description is not None:
                row[indices['description']].value = description
            if total_documents is not None:
                row[indices['total_documents']].value = str(total_documents)
            if num_files_downloaded is not None:
                row[indices['num_files_downloaded']].value = str(num_files_downloaded)
            if client_folder_path is not None:
                row[indices['client_folder_path']].value = client_folder_path
            
            # Update in client_list
            for client_info in self.client_list:
                if client_info['row_index'] == row_index:
                    if status is not None:
                        client_info['status_cell'].value = status
                    if description is not None:
                        client_info['description_cell'].value = description
                    if total_documents is not None:
                        client_info['total_documents_cell'].value = str(total_documents)
                    if num_files_downloaded is not None:
                        client_info['num_files_downloaded_cell'].value = str(num_files_downloaded)
                    if client_folder_path is not None:
                        client_info['client_folder_path_cell'].value = client_folder_path
                    break
            
        except (ExcelHeaderError, ExcelOperationError):
            raise
        except Exception as e:
            error_msg = f"Error updating client row {row_index}: {str(e)}"
            raise ExcelOperationError(error_msg) from e
    
    def update_document_row(self, row_index, download_status=None, download_desc=None,
                           file_name=None, file_path=None, folder_category=None,
                           download_time=None):
        """
        Update data to existing row by row_index in document_list_sheet
        
        Args:
            row_index (int): Row index to update
            download_status (str): Download status (optional)
            download_desc (str): Download description (optional)
            file_name (str): File name (optional)
            file_path (str): File path (optional)
            folder_category (str): Folder category (optional)
            download_time (str): Download time (optional)
            
        Raises:
            ExcelHeaderError: If header cannot be read
            ExcelOperationError: If there is an error updating document row
        """
        if not row_index:
            error_msg = "Row index cannot be empty"
            raise ExcelOperationError(error_msg)
        
        try:
            indices = self.get_document_header_indices()
            
            row = self.document_list_sheet[row_index]
            
            if download_status is not None:
                row[indices['download_status']].value = download_status
            if download_desc is not None:
                row[indices['download_desc']].value = download_desc
            if file_name is not None:
                row[indices['file_name']].value = file_name
            if file_path is not None:
                row[indices['file_path']].value = file_path
            if folder_category is not None:
                row[indices['folder_category']].value = folder_category
            if download_time is not None:
                row[indices['download_time']].value = download_time
            
            # Update in document_list
            for doc_info in self.document_list:
                if doc_info['row_index'] == row_index:
                    if download_status is not None:
                        doc_info['download_status'] = download_status
                    if download_desc is not None:
                        doc_info['download_desc'] = download_desc
                    if file_name is not None:
                        doc_info['file_name'] = file_name
                    if file_path is not None:
                        doc_info['file_path'] = file_path
                    if folder_category is not None:
                        doc_info['folder_category'] = folder_category
                    if download_time is not None:
                        doc_info['download_time'] = download_time
                    break
            
        except (ExcelHeaderError, ExcelOperationError):
            raise
        except Exception as e:
            error_msg = f"Error updating document row {row_index}: {str(e)}"
            raise ExcelOperationError(error_msg) from e
    
    def add_document_row(self, document_id, client_name, client_number,
                        file_name="", file_section="", document_type="",
                        description="", year="", document_date="",
                        file_size="", file_type="", folder_category=""):
        """
        Add new row to document_list sheet
        Before adding, check if document already exists in document_list
        If exists then update, if not then add new
        
        Args:
            document_id (str): Document ID
            client_name (str): Client Name
            client_number (str): Client Number
            file_name (str): File Name
            file_section (str): File Section
            document_type (str): Document Type
            description (str): Description
            year (str): Year
            document_date (str): Document Date
            file_size (str): File Size
            file_type (str): File Type
            folder_category (str): Folder Category
            
        Returns:
            int: Row index of row (new or updated)
            
        Raises:
            ExcelHeaderError: If header cannot be read
            ExcelOperationError: If there is an error adding document row
        """
        try:
            # Check if document already exists in document_list
            existing_row_index = self.get_document_row_index(client_name, client_number, document_id)
            
            if existing_row_index:
                # Already exists, perform update
                logger.debug(f"Document {document_id} already exists, updating row {existing_row_index}")
                # Update basic information (do not update download status)
                indices = self.get_document_header_indices()
                row = self.document_list_sheet[existing_row_index]
                row[indices['file_section']].value = file_section
                row[indices['doc_type']].value = document_type
                row[indices['description']].value = description
                row[indices['year']].value = year
                row[indices['doc_date']].value = document_date
                row[indices['file_size']].value = file_size
                row[indices['file_type']].value = file_type
                row[indices['folder_category']].value = folder_category
                if file_name:
                    row[indices['file_name']].value = file_name
                
                return existing_row_index
            else:
                # Not exists, add new
                indices = self.get_document_header_indices()
                
                # Create new row
                header_row = [cell.value for cell in self.document_list_sheet[1]]
                new_row = [None] * len(header_row)
                
                new_row[indices['download_status']] = ""
                new_row[indices['download_desc']] = ""
                new_row[indices['client_name']] = client_name
                new_row[indices['client_number']] = client_number
                new_row[indices['file_name']] = file_name
                new_row[indices['file_path']] = ""
                new_row[indices['folder_category']] = folder_category
                new_row[indices['file_section']] = file_section
                new_row[indices['doc_type']] = document_type
                new_row[indices['description']] = description
                new_row[indices['year']] = year
                new_row[indices['doc_date']] = document_date
                new_row[indices['file_size']] = file_size
                new_row[indices['doc_id']] = document_id
                new_row[indices['file_type']] = file_type
                new_row[indices['download_time']] = ""
                
                # Add row to sheet
                self.document_list_sheet.append(new_row)
                new_row_index = self.document_list_sheet.max_row
                
                # Add to document_list
                doc_info = {
                    'row_index': new_row_index,
                    'doc_id': document_id,
                    'client_name': client_name,
                    'client_number': client_number,
                    'download_status': "",
                    'download_desc': "",
                    'file_name': file_name,
                    'file_path': "",
                    'folder_category': folder_category,
                    'file_section': file_section,
                    'doc_type': document_type,
                    'description': description,
                    'year': year,
                    'doc_date': document_date,
                    'file_size': file_size,
                    'file_type': file_type,
                    'download_time': "",
                }
                self.document_list.append(doc_info)
                
                logger.debug(f"Added new document row: row {new_row_index}, doc_id: {document_id}")
                return new_row_index
                
        except (ExcelHeaderError, ExcelOperationError):
            raise
        except Exception as e:
            error_msg = f"Error adding document row: {str(e)}"
            raise ExcelOperationError(error_msg) from e
    
    def save_workbook(self):
        """
        Save workbook to Excel file
        
        Raises:
            ExcelSaveError: If file path is missing or there is an error saving
        """
        if not self.excel_file_path:
            error_msg = "No Excel file path to save"
            raise ExcelSaveError(error_msg)
        
        try:
            self.workbook.save(self.excel_file_path)
            logger.info(f"Saved Excel: {self.excel_file_path}")
        except Exception as e:
            error_msg = f"Error saving Excel: {str(e)}"
            raise ExcelSaveError(error_msg) from e
