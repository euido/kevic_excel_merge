import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QPushButton, QLabel, QVBoxLayout, QWidget
)
from openpyxl import load_workbook

class ExcelProcessor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("엑셀 처리 도구")
        self.setGeometry(100, 100, 400, 200)

        # Layout
        layout = QVBoxLayout()

        # Widgets
        self.label = QLabel("엑셀 파일을 선택하세요")
        layout.addWidget(self.label)

        self.select_button = QPushButton("파일 선택")
        self.select_button.clicked.connect(self.open_file_dialog)
        layout.addWidget(self.select_button)

        self.process_button = QPushButton("복사 및 병합 실행")
        self.process_button.clicked.connect(self.process_excel)
        layout.addWidget(self.process_button)

        # Container
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.file_path = ""

    def open_file_dialog(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, "엑셀 파일 선택", "", "Excel Files (*.xlsx);;All Files (*)", options=options
        )
        if file_path:
            self.file_path = file_path
            self.label.setText(f"선택된 파일: {os.path.basename(file_path)}")

    def process_excel(self):
        if not self.file_path:
            self.label.setText("먼저 엑셀 파일을 선택하세요.")
            return

        try:
            # Read Excel file
            df = pd.read_excel(self.file_path)

            # Copy columns A-E
            data = df.iloc[:, 0:5]

            # Remove duplicates based on columns A, B, C
            data = data.drop_duplicates(subset=[data.columns[0], data.columns[1], data.columns[2]])

            # Merge column D (index 3) into one cell, with line breaks
            data.iloc[:, 3] = data.groupby([data.columns[0], data.columns[1], data.columns[2]])[data.columns[3]] \
                .transform(lambda x: '\n'.join(x))

            # Copy to new columns H-L
            df_result = pd.DataFrame()
            df_result['H'] = data.iloc[:, 0]
            df_result['I'] = data.iloc[:, 1]
            df_result['J'] = data.iloc[:, 2]
            df_result['K'] = data.iloc[:, 3]

            # Save or update Excel file
            save_path, _ = QFileDialog.getSaveFileName(
                self, "저장할 파일 경로 선택", "", "Excel Files (*.xlsx);;All Files (*)"
            )
            if save_path:
                df_result.to_excel(save_path, index=False)
                self.label.setText("파일이 성공적으로 저장되었습니다!")
            else:
                self.label.setText("파일 저장이 취소되었습니다.")
        except Exception as e:
            self.label.setText(f"오류 발생: {e}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ExcelProcessor()
    ex.show()
    sys.exit(app.exec_())
