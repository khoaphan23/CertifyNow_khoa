# 📄 Tool Tạo Giấy Khen Tự Động

Công cụ tự động tạo giấy khen/chứng chỉ từ danh sách Excel và template Word, xuất trực tiếp ra PDF với nhiều tính năng nâng cao.

## ✨ Tính năng

### 🔥 Tính năng chính
- ✅ **Xử lý hàng loạt**: Tạo hàng trăm giấy khen chỉ trong vài phút
- 📊 **Đọc Excel thông minh**: Tự động phát hiện cấu trúc dữ liệu và lọc theo điều kiện
- 🎯 **Template linh hoạt**: Hỗ trợ placeholder ở bất kỳ vị trí nào trong Word template
- 📱 **Xuất PDF trực tiếp**: Không tạo file DOCX trung gian, tiết kiệm dung lượng
- 🕒 **Tự động đặt tên**: File được đặt tên theo STT và họ tên, có thể gộp thành 1 file PDF

### 🛠️ Tính năng nâng cao  
- 🔄 **Dual Engine**: python-docx (ưu tiên) + Word COM (fallback) đảm bảo thành công
- 🎨 **Giữ nguyên format**: Bảo toàn font chữ, màu sắc, định dạng của template gốc
- 📝 **Logging chi tiết**: Ghi log đầy đủ để debug và theo dõi quá trình xử lý
- ⚙️ **Cấu hình linh hoạt**: Tùy chỉnh mọi thông tin qua file config.ini
- 🗂️ **Quản lý file thông minh**: Tự động tạo thư mục, dọn dẹp file tạm

## 📁 Cấu trúc thư mục

```
CERTIFYNOW_KHOA/
├── 📄 README.md              # Hướng dẫn sử dụng
├── 📄 requirements.txt       # Danh sách thư viện cần thiết
├── 📄 main.py               # File chính để chạy
├── 📄 config.ini            # Cấu hình tập trung
├── 📄 __init__.py           # Package initialization
├── 📄 .gitignore            # Cấu hình Git ignore
├── 📂 input/                # Thư mục chứa file Excel
│   ├── 📊 danh_sach.xlsx
│   └── 📄 .gitkeep
├── 📂 templates/            # Thư mục chứa template Word
│   ├── 🎨 chung_chi_template.docx
│   └── 📄 .gitkeep
├── 📂 output/               # Thư mục chứa file PDF kết quả
├── 📂 temp/                 # Thư mục file tạm thời
├── 📂 logs/                 # Thư mục chứa log files
└── 📂 src/
    ├── 📄 __init__.py
    ├── 📂 certificate/
    │   └── 📄 generator.py       # Engine tạo giấy khen chính
    ├── 📂 io/
    │   └── 📄 file_handler.py    # Xử lý file và validation
    └── 📂 logging/
        └── 📄 logger_setup.py    # Thiết lập hệ thống logging
```

## 🚀 Cài đặt và sử dụng

### 1️⃣ Cài đặt các thư viện cần thiết
```bash
pip install -r requirements.txt
```

### 2️⃣ Windows (khuyến nghị cho PDF tối ưu)
```bash
pip install pywin32
```

### 3️⃣ Chuẩn bị dữ liệu

**File Excel** (đặt vào thư mục `input/`):
- Phải có các cột: `Họ và tên`, `Pháp danh`, `Năm sinh`, `Đơn vị`
- Dữ liệu bắt đầu từ hàng 5 (có thể thay đổi trong config)
- Hỗ trợ format .xlsx và .xls

**Template Word** (đặt vào thư mục `templates/`):
- File .docx chứa các placeholder theo format `<<Tên_placeholder>>`
- Có thể đặt placeholder ở paragraph, table, header, footer

### 4️⃣ Chạy tool
```bash
python main.py
```

### 5️⃣ Làm theo hướng dẫn
- Tool sẽ hiển thị cấu hình placeholder và danh sách người nhận
- Xác nhận trước khi bắt đầu tạo giấy khen
- File PDF sẽ được lưu trong thư mục `output/`

## 📋 Cấu hình chi tiết

### ⚙️ Chỉnh sửa file `config.ini`

**Thông tin cố định trên giấy khen:**
```ini
[CERTIFICATE]
issued_by = Ban Hướng Dẫn GĐPT TP Đà Nẵng
issued_at = Đà Nẵng  
issued_date = ngày 15 tháng 8 năm 2025
no_dharma_name = Không có
```

**Cấu hình đọc Excel:**
```ini
[EXCEL]
header_row = 5                 # Dữ liệu bắt đầu từ hàng 5
filter_column =                # Để trống = lấy tất cả dữ liệu
filter_value = 
```

**Cấu hình output:**
```ini
[OUTPUT]
create_individual_pdfs = true          # Tạo PDF riêng cho từng người
create_combined_pdf = true             # Gộp tất cả thành 1 file PDF
combined_pdf_name = Chung_chi_%Y%m%d_%H%M%S  # Tên file gộp
individual_pdf_format = %03d_%s        # Format: 001_Nguyen_Van_A.pdf
```

### 🕒 Placeholder thời gian
- `%Y` = năm 4 số (2025)
- `%m` = tháng 2 số (08) 
- `%d` = ngày 2 số (15)
- `%H` = giờ 2 số (14)
- `%M` = phút 2 số (30)
- `%S` = giây 2 số (25)

## 📝 Hướng dẫn tạo template

### 🎨 Template Word cần có các placeholder:

**Từ dữ liệu Excel (thay đổi theo từng người):**
- `<<Ho_va_ten>>` → Họ và tên
- `<<Phap_danh>>` → Pháp danh (hiển thị "Không có" nếu trống)
- `<<Nam_sinh>>` → Năm sinh
- `<<Don_vi>>` → Đơn vị

**Từ cấu hình (cố định cho tất cả):**
- `<<Do>>` → Cơ quan cấp
- `<<Tai>>` → Nơi cấp
- `<<Ngay>>` → Ngày cấp

### 📄 Ví dụ template Word:
```
                    CHỨNG CHỈ BẬC HƯỚNG THIỆN
                 GIA ĐÌNH PHẬT TỬ VIỆT NAM TP ĐÀ NẴNG

Chứng nhận: <<Ho_va_ten>>
Pháp danh: <<Phap_danh>>
Sinh năm: <<Nam_sinh>>
Hiện đang sinh hoạt tại: <<Don_vi>>

Đã trúng cách kỳ thi Bậc Hướng Thiện năm 2025
và được công nhận đạt yêu cầu.

                                        Do: <<Do>>
                                        Tại: <<Tai>>, <<Ngay>>
```

## 📊 Ví dụ file Excel

| STT | Họ và tên      | Pháp danh | Năm sinh | Đơn vị         | Ghi chú    |
|-----|---------------|-----------|----------|----------------|------------|
| 1   | Nguyễn Văn An | Minh Đức  | 1990     | GĐPT Linh Sơn  | Hải Châu   |
| 2   | Trần Thị Bình |           | 1985     | GĐPT Phổ Đà    | Thanh Khê  |
| 3   | Lê Văn Cường | Trí Tuệ   | 1992     | GĐPT Từ Ân     | Hải Châu   |

## 📱 Kết quả output

### 🗂️ File PDF riêng lẻ:
- `001_Nguyen_Van_An.pdf`
- `002_Tran_Thi_Binh.pdf`
- `003_Le_Van_Cuong.pdf`

### 📚 File PDF gộp:
- `Chung_chi_20250815_143025.pdf`

## 🔧 Xử lý lỗi và troubleshooting

### ⚡ Tool sử dụng dual engine:
1. **python-docx** (ưu tiên): Nhanh, ổn định, không cần Microsoft Word
2. **Word COM** (fallback): Chỉ trên Windows, cần Microsoft Word, xử lý format tốt hơn

### 📋 Kiểm tra khi gặp lỗi:
1. **Kiểm tra log**: Xem file log chi tiết trong thư mục `logs/`
2. **Kiểm tra placeholder**: Đảm bảo format `<<Tên_placeholder>>` đúng
3. **Kiểm tra Excel**: Các cột bắt buộc phải có dữ liệu
4. **Kiểm tra template**: File .docx không bị corrupt

### 🆘 Lỗi thường gặp:

**"Không tìm thấy file template":**
- Đặt file .docx vào thư mục `templates/`

**"Không tìm thấy file Excel":**
- Đặt file .xlsx/.xls vào thư mục `input/`

**"Không thể chuyển PDF":**
- Windows: Cài Microsoft Word hoặc LibreOffice
- Linux/Mac: Cài LibreOffice

## 🎯 Tính năng đặc biệt

### 🔍 **Smart Template Detection**
Tool tự động phát hiện và liệt kê tất cả placeholder trong template, kể cả khi bị tách thành nhiều đoạn trong Word.

### 🎨 **Format Preservation**  
Giữ nguyên định dạng font, màu sắc, kích thước chữ từ template gốc.

### 📊 **Progress Tracking**
Hiển thị tiến trình xử lý realtime với thanh progress và thống kê.

### 🧹 **Auto Cleanup**
Tự động dọn dẹp file tạm thời sau khi hoàn thành.

## 👥 Tác giả

**GĐPT Việt Nam - TP Đà Nẵng**  
Version: 1.0.0

---

*Tool được phát triển để hỗ trợ công tác tạo giấy khen tự động cho các hoạt động Phật pháp. Mọi góp ý xin liên hệ qua issues hoặc pull request.*