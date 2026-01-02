"""
Module xử lý mapping documents vào các category folders

Module này chứa logic để phân loại documents dựa trên File Section, 
Document Type và Description vào các thư mục category phù hợp.

Categories:
- _Permanent
- Accounting & Payroll
- Consulting & Special Projects
- Other
- Tax
"""

import logging

logger = logging.getLogger(__name__)


def get_document_category(file_section, document_type, description):
    """
    Xác định thư mục category dựa trên File Section, Document Type và Description
    
    Args:
        file_section (str): File Section của document (Bookkeeping hoặc Clientflow)
        document_type (str): Document Type của document
        description (str): Description của document
        
    Returns:
        str: Tên thư mục category
            Categories: _Permanent, Accounting & Payroll, Consulting & Special Projects, Other, Tax
    """
    try:
        # Chuẩn hóa input (lowercase, strip whitespace)
        file_section = (file_section or "").strip().lower()
        document_type = (document_type or "").strip().lower()
        description = (description or "").strip().lower()
        
        
        # Kiểm tra File Section và Document Type trước
        if file_section == "bookkeeping":
            return "Accounting & Payroll"
        
        elif file_section == "clientflow":
            return "Other"

        return "Other"
        
    except Exception as e:
        logger.warning(f"Lỗi khi xác định category: {str(e)}")
        return "Other"


def get_all_categories():
    """
    Trả về danh sách tất cả các categories có sẵn
    
    Returns:
        list: Danh sách các category names
    """
    return [
        "_Permanent",
        "Accounting & Payroll",
        "Consulting & Special Projects",
        "Other",
        "Tax"
    ]

