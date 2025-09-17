import os
from pathlib import Path
import shutil

def create_folders(folder_list):
    """Tạo các thư mục cần thiết"""
    for folder in folder_list:
        folder = Path(folder)
        if not folder.exists():
            folder.mkdir(parents=True, exist_ok=True)
            print(f"📁 Đã tạo thư mục: {folder}")

def validate_files(file_path, extensions):
    """Kiểm tra file có phải định dạng hợp lệ"""
    file_path = Path(file_path)
    if not file_path.exists():
        return False
    return file_path.suffix.lower() in extensions

def clean_temp_files(temp_folder):
    """Dọn dẹp thư mục tạm"""
    temp_folder = Path(temp_folder)
    if temp_folder.exists():
        shutil.rmtree(temp_folder)
        temp_folder.mkdir()
        print("🧹 Đã dọn dẹp thư mục temp")

def get_files_by_extension(folder, extensions):
    """Lấy danh sách file theo định dạng"""
    folder = Path(folder)
    files = []
    for ext in extensions:
        files.extend(folder.glob(f"*.{ext}"))
    return sorted(files)
