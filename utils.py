"""
Utils cho GOFILEROOM Downloader
Chứa các hàm tiện ích chung
"""

import os
import sys
import logging

logger = logging.getLogger(__name__)


def resource_path(relative_path):
    """
    Trả về đường dẫn tuyệt đối cho tài nguyên, hoạt động cho cả môi trường dev và PyInstaller.
    
    Args:
        relative_path (str): Đường dẫn tương đối
        
    Returns:
        str: Đường dẫn tuyệt đối
    """
    # Lấy đường dẫn thư mục chứa file thực thi (.exe) hoặc file script (.py)
    if getattr(sys, 'frozen', False):
        # Đang chạy dưới dạng file thực thi (.exe) đã đóng gói (PyInstaller)
        # sys.executable là đường dẫn tới file .exe
        base_path = os.path.dirname(sys.executable)
    else:
        # Đang chạy dưới dạng script Python (.py) trong môi trường dev
        # __file__ là đường dẫn tới script hiện tại
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)


def find_env_file():
    """
    Tìm file .env ở các vị trí có thể
    
    Returns:
        str: Đường dẫn đến file .env, None nếu không tìm thấy
    """
    # Thử tìm ở thư mục hiện tại
    env_path = os.path.join(os.getcwd(), '.env')
    if os.path.exists(env_path):
        return env_path
    
    # Thử tìm ở thư mục chứa file hiện tại
    current_dir = os.path.dirname(os.path.abspath(__file__))
    env_path = os.path.join(current_dir, '.env')
    if os.path.exists(env_path):
        return env_path
    
    return None


def load_env_config(raise_on_not_found=False):
    """
    Đọc cấu hình từ file .env
    
    Args:
        raise_on_not_found (bool): Nếu True, raise error khi không tìm thấy .env
        
    Returns:
        dict: Dictionary chứa các key-value từ .env, rỗng nếu không tìm thấy hoặc có lỗi
    """
    config_dict = {}
    
    try:
        env_path = find_env_file()
        
        if not env_path:
            error_msg = "File .env không tồn tại"
            if raise_on_not_found:
                logger.error(error_msg)
                raise FileNotFoundError(error_msg)
            else:
                logger.warning(error_msg)
                return config_dict
        
        with open(env_path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    key, value = line.split('=', 1)
                    config_dict[key.strip()] = value.strip()
        
        logger.info(f"Đã load config từ .env: {list(config_dict.keys())}")
        return config_dict
        
    except Exception as e:
        error_msg = f"Lỗi khi đọc file .env: {str(e)}"
        if raise_on_not_found:
            logger.error(error_msg)
            raise
        else:
            logger.warning(error_msg)
            return config_dict


def get_download_dir_from_env():
    """
    Lấy DOWNLOAD_DIR từ .env và xử lý đường dẫn
    
    Returns:
        str: Đường dẫn download directory, rỗng nếu không tìm thấy
    """
    config = load_env_config(raise_on_not_found=False)
    download_dir = config.get('DOWNLOAD_DIR', '')
    
    if not download_dir:
        return ""
    
    # Nếu là đường dẫn tuyệt đối thì dùng trực tiếp
    if os.path.isabs(download_dir):
        return download_dir
    
    # Nếu là đường dẫn tương đối thì dùng resource_path
    return resource_path(download_dir)



