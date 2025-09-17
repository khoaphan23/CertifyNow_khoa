# 🎓 Tool Tạo Giấy Khen Tự Động - GĐPT TP Đà Nẵng

Tool tự động tạo giấy khen/chứng chỉ từ danh sách Excel và phôi Word.

## ✨ Tính năng

- ✅ **Tự động điền thông tin**: Họ tên, pháp danh, năm sinh, đơn vị
- 📊 **Đọc danh sách Excel**: Hỗ trợ file .xlsx và .xls
- 📄 **Xuất PDF tự động**: Chuyển đổi từ Word sang PDF
- 🎯 **Xử lý hàng loạt**: Tạo nhiều giấy khen cùng lúc
- 📁 **Tổ chức thư mục**: Tự động tạo cấu trúc thư mục
- 📝 **Logging chi tiết**: Ghi nhận quá trình xử lý

## 📁 Cấu trúc thư mục

```
CERTIFYNOW_KHOA/
├── 📄 main.py               # File chính để chạy
├── 📄 config.ini           # File cấu hình
├── 📄 requirements.txt     # Thư viện cần thiết
├── 📄 README.md            # Hướng dẫn sử dụng
├── 📄 __init.py        
├── 📂 input/              # Chứa file Excel danh sách
├── 📂 templates/          # Chứa phôi giấy khen Word (.docx)
├── 📂 output/             # Kết quả sau khi xử lý
├── 📂 temp/               # File tạm thời
├── 📂 logs/               # File log
└── 📂 src/
    ├── 📂 certificate/    # Module xử lý giấy khen
    ├── 📂 io/            # Module xử lý file
    └── 📂 logging/       # Module logging
```

## 🚀 Cài đặt và sử dụng

### 1️⃣ Cài đặt thư viện
```bash
pip install -r requirements.txt
```

### 2️⃣ Chuẩn bị dữ liệu
- **File Excel**: Đặt file danh sách (format như mẫu) vào thư mục `input/`
- **Phôi Word**: Đặt file phôi giấy khen (.docx) vào thư mục `templates/`

### 3️⃣ Chạy chương trình
```bash
python main.py
```

### 4️⃣ Kết quả
- File Word và PDF sẽ được lưu trong thư mục `output/`
- File log trong thư mục `logs/`

## 📋 Format File Excel

File Excel cần có các cột sau (bắt đầu từ hàng 5):
- **Tt**: Số thứ tự
- **Họ và tên**: Tên người nhận giấy khen
- **Pháp danh**: Pháp danh (nếu có)
- **Năm sinh**: Năm sinh
- **Đơn vị**: Đơn vị/địa điểm
- **Điểm**: Điểm thi
- **Ghi chú**: Ghi chú (lọc "Hải Châu")

## 🎨 Format Phôi Word

Phôi Word (.docx) cần có các placeholder sau:
- `Chứng nhận :` → Sẽ được thay bằng họ tên
- `Pháp danh :` → Sẽ được thay bằng pháp danh
- `Sinh năm :` → Sẽ được thay bằng năm sinh
- `Hiện đang :` → Sẽ được thay bằng đơn vị

## ⚙️ Cấu hình

Chỉnh sửa file `config.ini` để tùy chỉnh:
- Đường dẫn thư mục
- Thông tin tổ chức
- Cấu hình Excel
- Cấu hình output

## 📝 Lưu ý

1. **Windows**: Cần cài Microsoft Word để chuyển đổi sang PDF
2. **Linux/Mac**: File sẽ giữ định dạng .docx
3. Tool sẽ tự động lọc những người có ghi chú "Hải Châu"

## 🛠️ Khắc phục sự cố

### Lỗi không đọc được Excel
- Kiểm tra file Excel có đúng định dạng
- Dữ liệu bắt đầu từ hàng 5
- Các cột có đúng tên

### Lỗi không tạo được PDF
- Windows: Cài đặt Microsoft Word
- Hoặc sử dụng file Word output

### Lỗi không tìm thấy phôi
- Đặt file .docx vào thư mục `templates/`
- Kiểm tra file không bị hỏng
