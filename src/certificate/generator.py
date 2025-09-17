import os
from pathlib import Path
from docx import Document
import logging
from datetime import datetime
import re
import sys

class CertificateGenerator:
    """Class xử lý tạo giấy khen - hỗ trợ textbox và shapes"""
    
    def __init__(self, template_path, logger=None, config=None):
        self.template_path = Path(template_path)
        self.logger = logger or logging.getLogger(__name__)
        self.config = config
        
        if config:
            self.issued_by = config.get('CERTIFICATE', 'issued_by', fallback='Ban Hướng Dẫn GĐPT')
            self.issued_at = config.get('CERTIFICATE', 'issued_at', fallback='Đà Nẵng')
            self.issued_date = config.get('CERTIFICATE', 'issued_date', fallback='').strip()
            self.no_dharma_name = config.get('CERTIFICATE', 'no_dharma_name', fallback='Không có')
            self.custom_placeholders = {}
            if config.has_section('PLACEHOLDERS'):
                for key, value in config.items('PLACEHOLDERS'):
                    placeholder_key = f'<<{key}>>'
                    self.custom_placeholders[placeholder_key] = value
        else:
            self.issued_by = 'Ban Hướng Dẫn GĐPT'
            self.issued_at = 'Đà Nẵng'
            self.issued_date = ''
            self.no_dharma_name = 'Không có'
            self.custom_placeholders = {}
        
        if not self.template_path.exists():
            raise FileNotFoundError(f"Không tìm thấy template: {template_path}")

    def create_certificate(self, ho_ten, phap_danh="", nam_sinh="", don_vi="", output_file=None):
        """Tạo giấy khen - ưu tiên dùng Word COM để xử lý textbox"""
        try:
            # Xử lý dữ liệu
            phap_danh_display = phap_danh.strip() if phap_danh.strip() else self.no_dharma_name
            
            if self.issued_date:
                current_date = self.issued_date
            else:
                current_date = datetime.now().strftime("ngày %d tháng %m năm %Y")
            
            if self.logger:
                self.logger.info(f"📝 Đang xử lý: {ho_ten}")
            
            # Tạo mapping đầy đủ
            replacements = {
                '<<Ho va ten>>': ho_ten,
                '<<Phap danh>>': phap_danh_display,
                '<<Nam sinh>>': str(nam_sinh) if nam_sinh else '',
                '<<Don vi>>': don_vi,
                '<<Do>>': self.issued_by,
                '<<Tai>>': self.issued_at,
                '<<Ngay>>': current_date,
            }
            
            if self.logger:
                self.logger.info("📄 Mapping sẽ sử dụng:")
                for k, v in replacements.items():
                    self.logger.info(f"  {k} → {v}")
            
            # Thử Word COM trước (có thể xử lý textbox)
            if sys.platform == "win32":
                success = self._use_word_com_advanced(replacements, output_file)
                if success:
                    return True
                else:
                    if self.logger:
                        self.logger.warning("⚠️ Word COM thất bại, thử phương pháp khác...")
            
            # Fallback: Dùng python-docx với xử lý run
            success = self._use_python_docx_advanced(replacements, output_file)
            return success
                
        except Exception as e:
            if self.logger:
                self.logger.error(f"❌ Lỗi tạo giấy khen cho {ho_ten}: {str(e)}")
            return False

    def _use_word_com_advanced(self, replacements, output_file):
        """Sử dụng Word COM với xử lý textbox và shapes"""
        try:
            import win32com.client
            from win32com.client import constants
        except ImportError:
            if self.logger:
                self.logger.warning("⚠️ Không có win32com")
            return False
        
        word = None
        try:
            if self.logger:
                self.logger.info("🔧 Đang sử dụng Word COM...")
            
            # Khởi tạo Word
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            
            # Mở document
            doc = word.Documents.Open(str(self.template_path.resolve()))
            
            total_replacements = 0
            
            # 1. Thay thế trong content thông thường
            for placeholder, replacement in replacements.items():
                # Main document
                find = doc.Content.Find
                find.ClearFormatting()
                find.Replacement.ClearFormatting()
                find.Text = placeholder
                find.Replacement.Text = str(replacement) if replacement else ''
                find.Forward = True
                find.Wrap = 1  # wdFindContinue
                find.Format = False
                find.MatchCase = False
                find.MatchWholeWord = False
                
                replaced = find.Execute(Replace=2)  # wdReplaceAll
                if replaced:
                    total_replacements += 1
                    if self.logger:
                        self.logger.info(f"   ✅ Main: {placeholder} → {replacement}")
            
            # 2. Thay thế trong headers/footers
            for section in doc.Sections:
                for header_type in [1, 2, 3]:  # Primary, FirstPage, EvenPages
                    try:
                        header = section.Headers(header_type)
                        if header.Exists:
                            for placeholder, replacement in replacements.items():
                                find = header.Range.Find
                                find.ClearFormatting()
                                find.Replacement.ClearFormatting()
                                find.Text = placeholder
                                find.Replacement.Text = str(replacement) if replacement else ''
                                replaced = find.Execute(Replace=2)
                                if replaced:
                                    total_replacements += 1
                    except:
                        pass
                
                for footer_type in [1, 2, 3]:
                    try:
                        footer = section.Footers(footer_type)
                        if footer.Exists:
                            for placeholder, replacement in replacements.items():
                                find = footer.Range.Find
                                find.ClearFormatting()
                                find.Replacement.ClearFormatting()
                                find.Text = placeholder
                                find.Replacement.Text = str(replacement) if replacement else ''
                                replaced = find.Execute(Replace=2)
                                if replaced:
                                    total_replacements += 1
                    except:
                        pass
            
            # 3. Thay thế trong shapes và textboxes
            try:
                for shape in doc.Shapes:
                    if shape.Type == 17:  # msoTextBox
                        try:
                            textbox_range = shape.TextFrame.TextRange
                            for placeholder, replacement in replacements.items():
                                if placeholder in textbox_range.Text:
                                    textbox_range.Text = textbox_range.Text.replace(
                                        placeholder, str(replacement) if replacement else ''
                                    )
                                    total_replacements += 1
                                    if self.logger:
                                        self.logger.info(f"   ✅ TextBox: {placeholder} → {replacement}")
                        except:
                            pass
            except:
                if self.logger:
                    self.logger.warning("⚠️ Không thể xử lý shapes")
            
            if self.logger:
                self.logger.info(f"📊 Word COM - Tổng thay thế: {total_replacements}")
            
            # Lưu file
            if total_replacements > 0 and output_file:
                doc.SaveAs2(str(Path(output_file).resolve()))
                doc.Close(False)
                word.Quit()
                if self.logger:
                    self.logger.info(f"✅ Tạo thành công bằng Word COM: {output_file.name}")
                return True
            else:
                doc.Close(False)
                word.Quit()
                return False
            
        except Exception as e:
            if self.logger:
                self.logger.error(f"❌ Lỗi Word COM: {e}")
            if word:
                try:
                    word.Quit()
                except:
                    pass
            return False

    def _use_python_docx_advanced(self, replacements, output_file):
        """Sử dụng python-docx với xử lý run-level"""
        try:
            if self.logger:
                self.logger.info("🔧 Đang sử dụng python-docx...")
            
            doc = Document(str(self.template_path))
            total_replacements = 0
            
            # Xử lý từng run riêng lẻ trong paragraphs
            for para_idx, para in enumerate(doc.paragraphs):
                # Ghép tất cả runs thành text đầy đủ
                full_text = ''
                for run in para.runs:
                    full_text += run.text
                
                if not full_text:
                    continue
                
                # Kiểm tra có placeholder không
                has_placeholder = any(placeholder in full_text for placeholder in replacements.keys())
                if not has_placeholder:
                    continue
                
                # Thay thế
                new_text = full_text
                for placeholder, replacement in replacements.items():
                    if placeholder in new_text:
                        new_text = new_text.replace(placeholder, str(replacement) if replacement else '')
                        total_replacements += 1
                        if self.logger:
                            self.logger.info(f"   ✅ Para[{para_idx}]: {placeholder} → {replacement}")
                
                # Cập nhật paragraph
                if new_text != full_text:
                    # Lưu format của run đầu tiên
                    original_style = None
                    if para.runs:
                        try:
                            first_run = para.runs[0]
                            original_style = {
                                'font_name': first_run.font.name,
                                'font_size': first_run.font.size,
                                'bold': first_run.font.bold,
                                'italic': first_run.font.italic,
                            }
                        except:
                            original_style = None
                    
                    # Xóa tất cả runs
                    for run in para.runs:
                        run.text = ''
                    
                    # Tạo run mới
                    new_run = para.add_run(new_text)
                    
                    # Áp dụng lại style
                    if original_style:
                        try:
                            if original_style.get('font_name'):
                                new_run.font.name = original_style['font_name']
                            if original_style.get('font_size'):
                                new_run.font.size = original_style['font_size']
                            if original_style.get('bold'):
                                new_run.font.bold = original_style['bold']
                            if original_style.get('italic'):
                                new_run.font.italic = original_style['italic']
                        except:
                            pass
            
            # Xử lý tables
            for table_idx, table in enumerate(doc.tables):
                for row_idx, row in enumerate(table.rows):
                    for cell_idx, cell in enumerate(row.cells):
                        for para_idx, para in enumerate(cell.paragraphs):
                            full_text = ''
                            for run in para.runs:
                                full_text += run.text
                            
                            if not full_text:
                                continue
                            
                            # Thay thế
                            new_text = full_text
                            for placeholder, replacement in replacements.items():
                                if placeholder in new_text:
                                    new_text = new_text.replace(placeholder, str(replacement) if replacement else '')
                                    total_replacements += 1
                                    if self.logger:
                                        self.logger.info(f"   ✅ Table[{table_idx}][{row_idx}][{cell_idx}][{para_idx}]: {placeholder} → {replacement}")
                            
                            # Cập nhật
                            if new_text != full_text:
                                for run in para.runs:
                                    run.text = ''
                                if para.runs:
                                    para.runs[0].text = new_text
                                else:
                                    para.add_run(new_text)
            
            if self.logger:
                self.logger.info(f"📊 python-docx - Tổng thay thế: {total_replacements}")
            
            # Lưu file
            if total_replacements > 0 and output_file:
                doc.save(str(output_file))
                if self.logger:
                    self.logger.info(f"✅ Tạo thành công bằng python-docx: {output_file.name}")
                return True
            else:
                if self.logger:
                    self.logger.error("❌ Không thể thay thế placeholder nào!")
                return False
                
        except Exception as e:
            if self.logger:
                self.logger.error(f"❌ Lỗi python-docx: {e}")
            return False

    def batch_create(self, data_list, output_folder):
        """Tạo nhiều giấy khen cùng lúc"""
        output_folder = Path(output_folder)
        output_folder.mkdir(parents=True, exist_ok=True)
        
        success_count = 0
        failed_list = []
        
        for idx, data in enumerate(data_list, 1):
            try:
                ho_ten = data.get('ho_ten', '')
                phap_danh = data.get('phap_danh', '')
                nam_sinh = data.get('nam_sinh', '')
                don_vi = data.get('don_vi', '')
                
                safe_name = ho_ten.replace(' ', '_').replace('/', '_').replace('\\', '_')
                output_file = output_folder / f"{idx:03d}_{safe_name}.docx"
                
                if self.create_certificate(ho_ten, phap_danh, nam_sinh, don_vi, output_file):
                    success_count += 1
                else:
                    failed_list.append(ho_ten)
                    
            except Exception as e:
                failed_list.append(data.get('ho_ten', f'Record {idx}'))
                if self.logger:
                    self.logger.error(f"❌ Lỗi xử lý {data.get('ho_ten', '')}: {str(e)}")
        
        return success_count, failed_list