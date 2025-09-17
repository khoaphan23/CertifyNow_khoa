import os
import sys
import pandas as pd
from docx import Document
from docx2pdf import convert
import shutil
from datetime import datetime
import configparser
from pathlib import Path

# Import các module từ src
from src.certificate.generator import CertificateGenerator
from src.io.file_handler import create_folders, validate_files
from src.logging.logger_setup import setup_logger

def load_config():
    """Đọc cấu hình từ file config.ini"""
    config = configparser.ConfigParser()
    config_file = 'config.ini'
    
    if os.path.exists(config_file):
        config.read(config_file, encoding='utf-8')
        print("📃 Đã tải cấu hình từ config.ini")
    else:
        print("⚠️ Không tìm thấy config.ini, sử dụng cấu hình mặc định")
    
    return config

def safe_str(value):
    """Chuyển đổi giá trị sang string an toàn"""
    if pd.isna(value) or value is None:
        return ""
    if isinstance(value, (int, float)):
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value)
    return str(value).strip()

def display_config_info(config):
    """Hiển thị thông tin cấu hình placeholder"""
    print("\n📋 THÔNG TIN CẤU HÌNH PLACEHOLDER:")
    print("-" * 70)
    print("Từ Excel (dữ liệu người nhận):")
    print("  • <<Họ_và_tên>> → Họ và tên")
    print("  • <<Pháp_danh>> → Pháp danh (nếu trống sẽ hiển thị 'Không có')")
    print("  • <<Năm_sinh>> → Năm sinh")
    print("  • <<Đơn_vị>> → Đơn vị")
    
    print("\nTừ Config (có thể chỉnh sửa trong config.ini):")
    issued_by = config.get('CERTIFICATE', 'issued_by', fallback='Ban Hướng Dẫn GĐPT')
    issued_at = config.get('CERTIFICATE', 'issued_at', fallback='Đà Nẵng')
    issued_date = config.get('CERTIFICATE', 'issued_date', fallback='')
    
    print(f"  • <<Do>> → {issued_by}")
    print(f"  • <<Tai>> → {issued_at}")
    print(f"  • <<Ngay>> → {issued_date if issued_date else 'Ngày hiện tại (tự động)'}")
    
    # Hiển thị placeholder tùy chỉnh nếu có
    if config.has_section('PLACEHOLDERS'):
        print("\nPlaceholder tùy chỉnh:")
        for key, value in config.items('PLACEHOLDERS'):
            print(f"  • <<{key}>> → {value}")
    
    print("-" * 70)
    print("📄 HƯỚNG DẪN SỬ DỤNG TRONG WORD TEMPLATE:")
    print("• Đặt các placeholder trên vào file Word template (*.docx)")
    print("• Ví dụ trong Word: 'Chứng nhận: <<Họ_và_tên>>'")
    print("• Ví dụ trong Word: 'Pháp danh: <<Pháp_danh>>'")
    print("• Có thể đặt ở bất kỳ đâu: paragraph, table, header, footer")
    print("-" * 70)

def main():
    """Hàm chính của chương trình"""
    
    # Khởi tạo logger
    logger = setup_logger("CertificateGenerator", "INFO", True)
    
    print("=" * 70)
    print("🎓 TOOL TẠO GIẤY KHEN TỰ ĐỘNG")
    print("   Gia Đình Phật Tử Việt Nam - TP Đà Nẵng")
    print("   📝 Sử dụng placeholder format: <<Tên_placeholder>>")
    print("=" * 70)
    
    # Đọc cấu hình
    config = load_config()
    
    # Hiển thị thông tin cấu hình
    display_config_info(config)
    
    # Thiết lập đường dẫn
    base_dir = Path.cwd()
    input_folder = base_dir / "input"
    output_folder = base_dir / "output"
    template_folder = base_dir / "templates"
    temp_folder = base_dir / "temp"
    
    # Tạo các thư mục cần thiết
    create_folders([input_folder, output_folder, template_folder, temp_folder])
    
    # Kiểm tra file template
    template_files = list(template_folder.glob("*.docx"))
    if not template_files:
        logger.error("❌ Không tìm thấy file phôi giấy khen (.docx) trong thư mục templates!")
        print("\n💡 Hướng dẫn:")
        print("1. Đặt file phôi giấy khen (định dạng .docx) vào thư mục 'templates'")
        print("2. File phôi cần chứa các placeholder:")
        print("   - <<Họ_và_tên>>, <<Pháp_danh>>, <<Năm_sinh>>, <<Đơn_vị>>")
        print("   - <<Do>>, <<Tai>>, <<Ngay>>")
        print("3. Chạy lại chương trình")
        return
    
    template_file = template_files[0]
    logger.info(f"📄 Sử dụng phôi: {template_file.name}")
    
    # Kiểm tra file Excel
    excel_files = list(input_folder.glob("*.xlsx")) + list(input_folder.glob("*.xls"))
    if not excel_files:
        logger.error("❌ Không tìm thấy file danh sách Excel trong thư mục input!")
        print("\n💡 Hướng dẫn:")
        print("1. Đặt file Excel chứa danh sách vào thư mục 'input'")
        print("2. File Excel cần có các cột: Họ và tên, Pháp danh, Năm sinh, Đơn vị")
        print("3. Chạy lại chương trình")
        return
    
    excel_file = excel_files[0]
    logger.info(f"📊 Đọc danh sách từ: {excel_file.name}")
    
    try:
        # Đọc dữ liệu từ Excel
        header_row = config.getint('EXCEL', 'header_row', fallback=5) - 1  # Convert to 0-based index
        df = pd.read_excel(excel_file, header=header_row)
        
        # Lọc và làm sạch dữ liệu
        df = df.dropna(subset=['Họ và tên'])  # Bỏ các hàng không có tên
        
        # Đổi tên cột cho dễ xử lý
        column_mapping = {
            'Tt': 'STT',
            'Họ và tên': 'HoTen',
            'Pháp danh': 'PhapDanh', 
            'Năm sinh': 'NamSinh',
            'Đơn vị': 'DonVi',
            'Điểm': 'Diem',
            'Ghi chú': 'GhiChu'
        }
        
        # Chỉ đổi tên các cột tồn tại
        existing_columns = {k: v for k, v in column_mapping.items() if k in df.columns}
        df = df.rename(columns=existing_columns)
        
        # Kiểm tra cột filter (nếu có cấu hình)
        filter_column = config.get('EXCEL', 'filter_column', fallback='')
        filter_value = config.get('EXCEL', 'filter_value', fallback='')
        
        if filter_column and filter_value and 'GhiChu' in df.columns:
            df_filtered = df[df['GhiChu'] == filter_value]
            if len(df_filtered) > 0:
                df = df_filtered
                logger.info(f"🔍 Đã lọc theo điều kiện: {filter_column} = {filter_value}")
        
        total_records = len(df)
        logger.info(f"📋 Tìm thấy {total_records} người trong danh sách")
        
        if total_records == 0:
            logger.error("❌ Không có dữ liệu hợp lệ để xử lý!")
            return
        
        # Hiển thị danh sách
        print("\n📋 DANH SÁCH NGƯỜI NHẬN GIẤY KHEN:")
        print("-" * 80)
        print(f"{'STT':>4} | {'Họ và tên':25} | {'Pháp danh':15} | {'Năm sinh':8} | {'Đơn vị'}")
        print("-" * 80)
        
        for idx, row in df.iterrows():
            try:
                stt = safe_str(row.get('STT', idx+1))
                ho_ten = safe_str(row['HoTen'])
                phap_danh = safe_str(row.get('PhapDanh', ''))
                nam_sinh = safe_str(row.get('NamSinh', ''))
                don_vi = safe_str(row.get('DonVi', ''))
                
                # Format an toàn cho việc hiển thị
                stt_num = int(float(stt)) if stt else (idx + 1)
                print(f"{stt_num:4d} | {ho_ten:25} | {phap_danh:15} | {nam_sinh:8} | {don_vi}")
                
            except Exception as e:
                logger.warning(f"Lỗi hiển thị dòng {idx}: {str(e)}")
        
        print("-" * 80)
        
        # Hỏi về việc chỉnh sửa config
        edit_config = input("Bạn có muốn dừng lại để chỉnh 'config.ini'? (y/N): ").strip().lower()
        if edit_config in ['y', 'yes']:
            print("➡️ Hãy mở file 'config.ini', chỉnh xong chạy lại chương trình.")
            return

        # Xác nhận tạo giấy khen
        confirm = input(f"\n❓ Tiến hành tạo {total_records} giấy khen? (y/N): ").strip().lower()
        if confirm not in ['y', 'yes']:
            print("❌ Đã hủy!")
            return

        # Khởi tạo generator với config
        generator = CertificateGenerator(template_file, logger, config)

        # Thư mục tạm và output
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        temp_folder.mkdir(exist_ok=True)

        temp_word_files = []
        pdf_files = []
        success_count = 0

        print("\n📄 Đang xử lý...")
        print("-" * 60)

        for idx, row in df.iterrows():
            try:
                stt_raw = row.get('STT', idx+1)
                stt = int(float(stt_raw)) if pd.notna(stt_raw) else (idx + 1)

                ho_ten = safe_str(row['HoTen'])
                phap_danh = safe_str(row.get('PhapDanh', ''))
                nam_sinh = safe_str(row.get('NamSinh', ''))
                don_vi = safe_str(row.get('DonVi', ''))

                safe_filename = ho_ten.replace(' ', '_').replace('/', '_').replace('\\', '_')
                temp_word_path = temp_folder / f"{stt:03d}_{safe_filename}.docx"

                print(f"  [{stt:2d}/{total_records}] Đang xử lý: {ho_ten}... ", end='')

                ok = generator.create_certificate(
                    ho_ten=ho_ten,
                    phap_danh=phap_danh,
                    nam_sinh=nam_sinh,
                    don_vi=don_vi,
                    output_file=temp_word_path
                )

                if ok:
                    temp_word_files.append(temp_word_path)
                    success_count += 1
                    print("✅")
                else:
                    print("❌")
            except Exception as e:
                logger.error(f"Lỗi xử lý {row.get('HoTen', 'Unknown')}: {str(e)}")
                print("❌")

        print("-" * 60)

        # Chuyển đổi sang PDF (Windows có MS Word)
        if temp_word_files:
            print(f"\n📄 Đang chuyển đổi {len(temp_word_files)} file sang PDF...")
            for word_file in temp_word_files:
                try:
                    pdf_out = output_folder / f"{word_file.stem}.pdf"
                    if sys.platform == "win32":
                        try:
                            convert(str(word_file), str(pdf_out))
                            pdf_files.append(pdf_out)
                        except Exception:
                            # Nếu không chuyển được, copy DOCX ra output
                            shutil.copy2(word_file, output_folder / word_file.name)
                            logger.warning(f"Không thể chuyển PDF, giữ DOCX: {word_file.name}")
                    else:
                        shutil.copy2(word_file, output_folder / word_file.name)
                except Exception as e:
                    logger.error(f"Lỗi chuyển đổi {word_file.name}: {str(e)}")

            # Gộp PDF nếu có
            try:
                from PyPDF2 import PdfMerger
                if pdf_files:
                    merger = PdfMerger()
                    for pdf in sorted(pdf_files):
                        merger.append(str(pdf))
                    combined_pdf = output_folder / f"GiayKhen_TongHop_{timestamp}.pdf"
                    merger.write(str(combined_pdf))
                    merger.close()
                    logger.info(f"✅ Đã gộp PDF: {combined_pdf.name}")
            except ImportError:
                logger.info("🔌 Cài đặt PyPDF2 để gộp các file PDF")
            except Exception as e:
                logger.warning(f"Không thể gộp PDF: {str(e)}")

        # Dọn dẹp
        print("\n🧹 Dọn dẹp file tạm...")
        for p in temp_word_files:
            try:
                p.unlink()
            except Exception:
                pass

        # Kết quả
        print("\n" + "=" * 60)
        print("✅ HOÀN THÀNH!")
        print(f"📊 Đã tạo: {success_count}/{total_records} giấy khen")
        print(f"📁 Thư mục kết quả: {output_folder}")
        print("=" * 60)

        # Mở thư mục output
        open_folder = input("\n🗂️ Mở thư mục kết quả? (y/N): ").strip().lower()
        if open_folder in ['y', 'yes']:
            try:
                import platform
                system = platform.system().lower()
                
                if system == "windows":
                    os.startfile(str(output_folder))
                elif system == "darwin":  # macOS
                    import subprocess
                    subprocess.Popen(["open", str(output_folder)])
                else:  # Linux và các Unix-like systems
                    import subprocess
                    subprocess.Popen(["xdg-open", str(output_folder)])
                    
                print(f"📂 Đã mở thư mục: {output_folder}")
            except Exception as e:
                print(f"⚠️ Không thể mở thư mục tự động: {str(e)}")
                print(f"📁 Vui lòng mở thủ công: {output_folder}")

    except Exception as e:
        logger.error(f"Lỗi chính: {str(e)}")
        print(f"\n❌ Đã xảy ra lỗi: {str(e)}")

if __name__ == "__main__":
    main()