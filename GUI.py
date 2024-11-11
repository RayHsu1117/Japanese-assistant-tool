from PyQt5.QtWidgets import QWidget, QVBoxLayout, QPushButton, QFileDialog, QTextEdit, QLabel
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

class CsvToExcelApp(QWidget):
    def __init__(self, import_function):
        super().__init__()

        # 設定視窗
        self.setWindowTitle("CSV to Excel Importer")
        self.setGeometry(100, 100, 600, 400)

        # 將匯入函數存入以後使用
        self.import_function = import_function

        # 設定布局
        layout = QVBoxLayout()

        # 標籤與按鈕
        self.root_label = QLabel("選擇包含資料夾的主目錄")
        self.root_label.setAlignment(Qt.AlignCenter)
        self.set_label_font(self.root_label)  # 設定字體大小
        layout.addWidget(self.root_label)

        self.select_root_button = QPushButton("選擇主目錄")
        self.set_button_font(self.select_root_button)  # 設定字體大小
        self.select_root_button.clicked.connect(self.select_root_folder)
        layout.addWidget(self.select_root_button)

        self.excel_label = QLabel("選擇 Excel 檔案（或自動創建）")
        self.excel_label.setAlignment(Qt.AlignCenter)
        self.set_label_font(self.excel_label)  # 設定字體大小
        layout.addWidget(self.excel_label)

        self.select_excel_button = QPushButton("選擇 Excel 檔案")
        self.set_button_font(self.select_excel_button)  # 設定字體大小
        self.select_excel_button.clicked.connect(self.select_excel_file)
        layout.addWidget(self.select_excel_button)

        # 啟動匯入按鈕
        self.start_button = QPushButton("開始匯入 CSV 到 Excel")
        self.set_button_font(self.start_button)  # 設定字體大小
        self.start_button.clicked.connect(self.start_import)
        layout.addWidget(self.start_button)

        # 結果顯示區
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        layout.addWidget(self.result_text)

        self.setLayout(layout)
        self.root_folder = ""
        self.excel_file = ""

        # 設定結果顯示字體大小
        font = QFont()
        font.setPointSize(16)  # 設置字體大小為 16
        self.result_text.setFont(font)

    def set_label_font(self, label):
        font = QFont()
        font.setPointSize(14)  # 設定 QLabel 字體大小
        label.setFont(font)

    def set_button_font(self, button):
        font = QFont()
        font.setPointSize(14)  # 設定 QPushButton 字體大小
        button.setFont(font)

    def select_root_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "選擇主目錄")
        if folder:
            self.root_folder = folder
            self.root_label.setText(f"主目錄: {folder}")

    def select_excel_file(self):
        file, _ = QFileDialog.getSaveFileName(self, "選擇或創建 Excel 檔案", "", "Excel Files (*.xlsx)")
        if file:
            self.excel_file = file
            self.excel_label.setText(f"Excel 檔案: {file}")

    def start_import(self):
        if self.root_folder and self.excel_file:
            self.import_function(self.root_folder, self.excel_file, self.display_result)
        else:
            self.display_result("請先選擇主目錄和 Excel 檔案")

    def display_result(self, message):
        # 直接將訊息顯示於 QTextEdit，並且確保訊息中有換行符（\n）
        self.result_text.append(message)
