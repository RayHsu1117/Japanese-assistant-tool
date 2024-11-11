from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog, QTextEdit, QLabel
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QApplication , QMainWindow
from csv2excel import import_csv_to_excel  # 匯入 csv2excel.py 中的函數
from seat import generate_seating_chart 
from seat import generate_seating_chart_for_all_csvs  # 匯入生成座位表的功能
import sys

class CsvToExcelApp(QWidget):
    def __init__(self, import_function):
        super().__init__()

        # 設定視窗
        self.setWindowTitle("CSV to Excel Importer")
        self.setGeometry(100, 100, 600, 400)

        # 將匯入函數存入以後使用
        self.import_function = import_function

# 設定布局
        main_layout = QVBoxLayout()

        # 設置主目錄選擇區域
        root_layout = QVBoxLayout()
        self.root_label = QLabel("選擇包含資料夾的主目錄")
        self.root_label.setAlignment(Qt.AlignCenter)
        self.set_label_font(self.root_label)
        root_layout.addWidget(self.root_label)

        self.select_root_button = QPushButton("選擇資料夾")
        self.set_button_font(self.select_root_button)
        self.select_root_button.clicked.connect(self.select_root_folder)
        root_layout.addWidget(self.select_root_button)
        main_layout.addLayout(root_layout)

              # 選擇 Excel 和 Word 檔案的區域，並排列為水平佈局
        file_selection_layout = QHBoxLayout()

        # 選擇 Excel 檔案的布局
        excel_layout = QVBoxLayout()
        self.excel_label = QLabel("選擇 Excel 檔案")
        self.excel_label.setAlignment(Qt.AlignCenter)
        self.set_label_font(self.excel_label)
        excel_layout.addWidget(self.excel_label)

        self.select_excel_button = QPushButton("選取 Excel")
        self.set_button_font(self.select_excel_button)
        self.select_excel_button.clicked.connect(self.select_excel_file)
        excel_layout.addWidget(self.select_excel_button)

        file_selection_layout.addLayout(excel_layout)

        # 選擇 Word 檔案的布局
        word_layout = QVBoxLayout()
        self.word_label = QLabel("選擇 Word 資料夾")
        self.word_label.setAlignment(Qt.AlignCenter)
        self.set_label_font(self.word_label)
        word_layout.addWidget(self.word_label)

        self.select_word_button = QPushButton("選取 Word 資料夾")
        self.set_button_font(self.select_word_button)
        self.select_word_button.clicked.connect(self.select_word_folder)
        word_layout.addWidget(self.select_word_button)

        file_selection_layout.addLayout(word_layout)
        main_layout.addLayout(file_selection_layout)
        # 設置功能區域
        function_layout = QHBoxLayout()

        self.start_button = QPushButton("匯入 Excel")
        self.set_button_font(self.start_button)
        self.start_button.clicked.connect(self.set_import_function_to_csv)
        function_layout.addWidget(self.start_button)

        self.generate_seating_button = QPushButton("產生座位表")
        self.set_button_font(self.generate_seating_button)
        self.generate_seating_button.clicked.connect(self.set_import_function_to_seating)
        function_layout.addWidget(self.generate_seating_button)

        main_layout.addLayout(function_layout)
        # 結果顯示區
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        main_layout.addWidget(self.result_text)

        self.setLayout(main_layout)
        self.root_folder = ""
        self.excel_file = ""  # 用於存儲 Excel 檔案路徑
        self.word_folder = ""  # 用於存儲 Word 檔案路徑
        self.import_function = None  # 初始化 import_function 為 None


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

    def select_word_folder(self):
        word_folder = QFileDialog.getExistingDirectory(self, "選擇 Word 資料夾")
        if word_folder:
            self.word_folder = word_folder
            self.word_label.setText(f"主目錄: {word_folder}")
        
    def set_import_function_to_csv(self):
        self.import_function = import_csv_to_excel
        if self.root_folder and self.excel_file:
            self.import_function(self.root_folder, self.excel_file, self.display_result)
        else:
            self.display_result("請先選擇主目錄和 Excel 檔案")

    def set_import_function_to_seating(self):
        """設置匯入功能為座位表生成"""
        self.import_function = generate_seating_chart_for_all_csvs
        if self.root_folder and self.word_folder:
            self.display_result("已設置為生成座位表。")
            self.import_function(self.root_folder, self.word_folder, self.display_result)
        else:
            self.display_result("請先選擇主目錄和 Word 資料夾")

    def display_result(self, message):
        # 直接將訊息顯示於 QTextEdit，並且確保訊息中有換行符（\n）
        self.result_text.append(message)
    
    def generate_seating_chart(self):
        if self.root_folder:
            # 呼叫 seat.py 中的函數來生成座位表
            results = generate_seating_chart_for_all_csvs(self.root_folder)
            for result in results:
                self.display_result(result)
        else:
            self.display_result("請先選擇資料夾來生成座位表")

# 啟動應用程式
if __name__ == "__main__":
    # 必須在創建任何 PyQt5 元件之前創建 QApplication
    app = QApplication(sys.argv)
    
    # 使用從 csv2excel.py 匯入的函數
    window = CsvToExcelApp(None)
    window.show()

    sys.exit(app.exec_())