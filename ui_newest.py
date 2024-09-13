"""
Author: tienva@bidv.com.vn
Version: 1
Date: 12/09/2024
"""
"""
    When a file (like a DOCX) is open in another program (such as Microsoft Word),
    it is typically locked by the operating system, preventing other applications from modifying it.
    This is why your labeling function fails when the DOCX file is open in Windows
"""


import sys
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QLineEdit, QPushButton, QTextEdit, QFileDialog,
                             QMessageBox, QDesktopWidget, QMenu)
from main_newest import scan_file, label_docx_file, label_xlsx_file_footer, classify_document_with_multiple_rules, define_rules, is_file_locked
import json
from docx import Document
from docx.shared import Pt
from openpyxl import load_workbook




class LabelingApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()


    def initUI(self):
        """Initialize the user interface"""
        self.setWindowTitle('Label tool')
        self.setGeometry(100, 100, 600, 400)
        self.center()


        # Main layout
        main_layout = QVBoxLayout()


        # File info layout (file selection)
        file_info_layout = QHBoxLayout()


        # Label for file path
        self.file_path_label = QLabel('File path')
        file_info_layout.addWidget(self.file_path_label)


        # Text field to display the selected file path
        self.file_path_edit = QLineEdit()
        file_info_layout.addWidget(self.file_path_edit)


        # Browse button to select a file
        self.browse_button = QPushButton('Browse')
        self.browse_button.clicked.connect(self.browse_file)
        file_info_layout.addWidget(self.browse_button)


        main_layout.addLayout(file_info_layout)


        # File operations layout (buttons for scanning and labeling)
        file_ops_layout = QHBoxLayout()


        # Button to scan the file
        self.scan_button = QPushButton('Scan')
        self.scan_button.clicked.connect(self.handle_scan_file)
        file_ops_layout.addWidget(self.scan_button)


        # Button to label the file
        self.label_button = QPushButton('Label file')
        self.label_button.clicked.connect(self.handle_label_file)
        file_ops_layout.addWidget(self.label_button)


        # Button for editing label (with options Bí mật, Nội bộ, Công khai)
        self.edit_label_button = QPushButton('Edit Label')
        self.edit_label_menu = QMenu(self)


        # Add options to the Edit Label menu
        self.edit_label_menu.addAction("Bí mật", lambda: self.set_label("Bí mật"))
        self.edit_label_menu.addAction("Nội bộ", lambda: self.set_label("Nội bộ"))
        self.edit_label_menu.addAction("Công khai", lambda: self.set_label("Công khai"))


        # Connect the button to show the menu when clicked
        self.edit_label_button.setMenu(self.edit_label_menu)
        file_ops_layout.addWidget(self.edit_label_button)


        # Button to show file label info
        self.file_info_button = QPushButton('Labeled file information')
        self.file_info_button.clicked.connect(self.handle_file_info)
        file_ops_layout.addWidget(self.file_info_button)


        main_layout.addLayout(file_ops_layout)


        # Text area to display file details or scan results
        self.file_details_text = QTextEdit()
        main_layout.addWidget(self.file_details_text)


        # Set the main layout for the window
        self.setLayout(main_layout)


    def center(self):
        """Centers the window on the screen"""
        screen = QDesktopWidget().screenGeometry()
        window = self.frameGeometry()
        window.moveCenter(screen.center())
        self.move(window.topLeft())


    def browse_file(self):
        """Opens a file dialog to browse and select a file"""
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, "Chọn file", "", "All Files (*)")
        if file_path:
            self.file_path_edit.setText(file_path)


    def handle_scan_file(self):
        """Handles the scan operation and displays the results."""
        file_path = self.file_path_edit.text()  # Get the file path from the UI
        if not file_path:
            self.file_details_text.setText("Vui lòng chọn một tệp trước khi scan.")
            return
       
        if is_file_locked(file_path):
            self.file_details_text.setText("File đang mở. Vui lòng đóng file trước khi quét tệp.")
            return
       
        results, message = scan_file(file_path)


        if message == "Không hỗ trợ định dạng tệp.":
            self.file_details_text.setText(message)
            return


        rules = define_rules()
        category, show_rule = classify_document_with_multiple_rules(results, rules)


        if category == "Confidential":
            # formatted_result = json.dumps(matched_keys, indent=4, ensure_ascii=False)
            self.file_details_text.setText(show_rule)
        else:
            self.file_details_text.setText("Tài liệu Nội bộ của BIDV. Cấm sao chép, in ấn dưới mọi hình thức!")


    def handle_label_file(self):
        """Handles the labeling operation for DOCX and XLSX files."""
        file_path = self.file_path_edit.text()


        if not file_path:
            self.file_details_text.setText("Vui lòng chọn một tệp trước khi label.")
            return


        if is_file_locked(file_path):
            self.file_details_text.setText("File đang mở. Vui lòng đóng file trước khi gán nhãn.")
            return


        if file_path.endswith('.docx'):
            results, message = scan_file(file_path)
            if results == "Không hỗ trợ định dạng tệp":
                self.file_details_text.setText(message)
                return


            rules = define_rules()
            category, show_rule = classify_document_with_multiple_rules(results, rules)
            if category:


                label_text = label_docx_file(file_path, category)


                self.file_details_text.setText(f"Đã gán nhãn '{label_text}' cho tài liệu.")
                msg_box = QMessageBox()
                msg_box.setIcon(QMessageBox.Information)
                msg_box.setText(f"Đã gán nhãn '{label_text}' cho file thành công.")
                msg_box.setWindowTitle("Thông báo")
                msg_box.exec_()


        elif file_path.endswith(".xlsx"):
            results, message = scan_file(file_path)
            if results == "Không hỗ trợ định dạng tệp":
                self.file_details_text.setText(message)
                return
            rules = define_rules()
            category, show_rule = classify_document_with_multiple_rules(results, rules)


            if category:
                label_text = label_xlsx_file_footer(file_path, category)


                self.file_details_text.setText(f"Đã gán nhãn '{label_text}' cho tài liệu.")
                msg_box = QMessageBox()
                msg_box.setIcon(QMessageBox.Information)
                msg_box.setText(f"Đã gán nhãn '{label_text}' cho file thành công.")
                msg_box.setWindowTitle("Thông báo")
                msg_box.exec_()


        else:
            self.file_details_text.setText("Chức năng gán nhãn chỉ hỗ trợ định dạng DOCX, XLSX, CSV.")


    def set_label(self, label_type):
        """Set the label type based on user selection and ask for confirmation"""
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Question)
        msg_box.setWindowTitle("Xác nhận")
        msg_box.setText(f"Bạn có chắc chắn muốn gán nhãn '{label_type}' cho file không?")
        msg_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        result = msg_box.exec_()


        if result == QMessageBox.Yes:
            self.file_details_text.setText(f"Đang gán nhãn: {label_type}")
            self.label_file_with_new_label(label_type)


    def label_file_with_new_label(self, label_type):
        """Label the file with the new label selected by the user"""
        file_path = self.file_path_edit.text()


        if not file_path:
            self.file_details_text.setText("Vui lòng chọn một tệp trước khi label.")
            return


        if is_file_locked(file_path):
            self.file_details_text.setText("File đang mở. Vui lòng đóng file trước khi gán nhãn.")
            return


        if file_path.endswith('.docx'):
            label_text = self.edit_label_docx_file(file_path, label_type)
            self.file_details_text.setText(f"Đã gán nhãn '{label_text}' cho file DOCX.")
        elif file_path.endswith('.xlsx'):
            label_text = self.edit_label_xlsx_file_footer(file_path, label_type)
            self.file_details_text.setText(f"Đã gán nhãn '{label_text}' cho file XLSX.")
        else:
            self.file_details_text.setText("Chỉ hỗ trợ gán nhãn cho file DOCX và XLSX.")


    def edit_label_docx_file(self, file_path, label_type):
        """Add or overwrite a label in the footer of a DOCX file based on user selection."""
        document = Document(file_path)


        if label_type == "Bí mật":
            label_text = "Tài liệu Mật của BIDV. Cấm sao chép, in ấn dưới mọi hình thức!"
        elif label_type == "Nội bộ":
            label_text = "Tài liệu Nội bộ của BIDV. Cấm sao chép, in ấn dưới mọi hình thức!"
        else:
            label_text = "Tài liệu Công khai của BIDV."


            # Clear existing footer text
        for section in document.sections:
            footer = section.footer
            for paragraph in footer.paragraphs:
                paragraph.clear()


            # Thêm đoạn văn mới trong footer để chứa nhãn và căn trái
            paragraph = footer.add_paragraph()
            run = paragraph.add_run(label_text)
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
            run.bold = False  # Đảm bảo văn bản không in đậm
            paragraph.alignment = 0  # Căn trái cho đoạn văn


            # # Thêm số trang vào footer (đặt ở vị trí chính giữa)
            # paragraph_page = footer.add_paragraph()
            # run_page = paragraph_page.add_run("Page ")
            # run_page.font.size = Pt(12)
            # run_page.font.name = 'Times New Roman'
            # run_page.bold = False  # Không in đậm


            # # Chèn mã trường để tự động hiển thị số trang
            # run_page.add_field('PAGE')
            # paragraph_page.alignment = 2  # Căn giữa cho số trang (giá trị 1 là căn giữa)


        document.save(file_path)
        return label_text






    def edit_label_xlsx_file_footer(self, file_path, label_type):
        """Add or overwrite a label in the footer of an XLSX file based on user selection."""
        workbook = load_workbook(file_path)
        sheet = workbook.active


        # Xác định văn bản của nhãn dựa trên loại nhãn
        if label_type == "Bí mật":
            label_text = "Tài liệu Mật của BIDV. Cấm sao chép, in ấn dưới mọi hình thức!"
        elif label_type == "Nội bộ":
            label_text = "Tài liệu Nội bộ của BIDV. Cấm sao chép, in ấn dưới mọi hình thức!"
        else:
            label_text = "Tài liệu Công khai của BIDV."


        # Gán nhãn vào phần footer của file Excel, căn trái và font Times New Roman cỡ 12
        sheet.oddFooter.left.text = label_text
        sheet.oddFooter.left.size = 12  # Cỡ chữ là 12
        sheet.oddFooter.left.font = "Times New Roman"


        # Giữ lại số trang ở vị trí chính giữa
        sheet.oddFooter.center.text = "Page &P"
        sheet.oddFooter.center.size = 12
        sheet.oddFooter.center.font = "Times New Roman"


        workbook.save(file_path)
        return label_text




    def handle_file_info(self):
        """Handles displaying the file path and label text information"""
        file_path = self.file_path_edit.text()
        if not file_path:
            self.file_details_text.setText("Vui lòng chọn một tệp trước khi hiển thị thông tin.")
            return


        if is_file_locked(file_path):
            self.file_details_text.setText("File đang mở. Vui lòng đóng file trước khi xem thông tin tệp.")
            return
       
        # Label the DOCX file based on scan results
        results, message = scan_file(file_path)
        rules = define_rules()  # Fetch defined rules
        category = classify_document_with_multiple_rules(results, rules)


        if file_path.endswith('.docx'):
            label_text = label_docx_file(file_path, category)
        elif file_path.endswith('.xlsx'):
            label_text = label_xlsx_file_footer(file_path, category)
        else:
            label_text = "Hiện tại chỉ hỗ trợ file DOCX, XLSX, CSV."


        # Create a dictionary with the file path and final label
        result_info = {
            "Path": file_path        
            }


        # Convert the dictionary to JSON format
        file_info = json.dumps(result_info, indent=4, ensure_ascii=False)


        # Display the JSON result in the text box
        self.file_details_text.setText(file_info)
        return file_info




if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = LabelingApp()
    ex.show()
    sys.exit(app.exec_())