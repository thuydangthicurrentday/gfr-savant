import os
from pathlib import Path

# Đường dẫn thư mục cần đọc
folder_path = r"C:\Users\gfr-admin\Savant Capital, LLC\Karbon - Documents"

# Kiểm tra thư mục có tồn tại không
if os.path.exists(folder_path):
    # Lấy danh sách các item trong thư mục
    items = os.listdir(folder_path)
    
    # Lọc chỉ lấy các thư mục (folder), không lấy file
    folders = [item for item in items if os.path.isdir(os.path.join(folder_path, item))]
    
    # Sắp xếp theo tên (tùy chọn)
    folders.sort()
    
    # In từng folder ra màn hình
    for folder in folders:
        print(folder)
else:
    print(f"Thư mục không tồn tại: {folder_path}")





