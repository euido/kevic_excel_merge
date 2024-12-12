import sys
import time
from PyQt5.QtWidgets import QApplication, QVBoxLayout, QPushButton, QLabel, QFileDialog, QWidget, QLineEdit, QHBoxLayout
from PyQt5.QtCore import QThread, pyqtSignal
from openpyxl import load_workbook


class ProcessThread(QThread):
    # Signal to communicate status messages
    progress = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, file_path, min_row, max_row):
        super().__init__()
        self.file_path = file_path
        self.min_row = min_row
        self.max_row = max_row

    def run(self):
        try:
            self.progress.emit("Loading Excel file...")
            workbook = load_workbook(self.file_path)
            first_sheet = workbook.active  # 첫 번째 시트를 가져옴
            
            if not first_sheet:
                self.progress.emit("No sheet found in the workbook.")
                self.finished.emit()
                return

            # 범위 내 B열 값 기반 시트 복사
            self.progress.emit("Processing B열 and creating new sheets...")
            for row in range(self.min_row, self.max_row + 1):  # 지정 범위만큼 반복
                b_value = first_sheet.cell(row=row, column=2).value  # B열 값 읽기
                if b_value:
                    new_sheet_name = str(b_value)  # B열 값을 시트 이름으로
                    if new_sheet_name in workbook.sheetnames:
                        self.progress.emit(f"Sheet '{new_sheet_name}' already exists, skipping.")
                        continue

                    # 첫 번째 시트를 복사
                    new_sheet = workbook.copy_worksheet(first_sheet)
                    new_sheet.title = new_sheet_name

            workbook.save(self.file_path)
            self.progress.emit("Sheets created and saved successfully.")
            self.finished.emit()
        except Exception as e:
            self.progress.emit(f"Error: {e}")
            self.finished.emit()


class MainApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel 시트 복사 및 B열 값 기반 이름")
        self.resize(400, 400)

        # UI 요소
        self.layout = QVBoxLayout()
        self.select_file_btn = QPushButton("Select Source Excel")
        self.min_row_label = QLabel("Min Row:")
        self.min_row_input = QLineEdit()
        self.max_row_label = QLabel("Max Row:")
        self.max_row_input = QLineEdit()
        self.process_btn = QPushButton("Process B열 and Create Sheets")
        self.status_label = QLabel("Select an Excel file to process.")

        # UI 구성
        self.layout.addWidget(self.select_file_btn)
        
        # Min/Max 입력 필드
        self.range_layout = QHBoxLayout()
        self.range_layout.addWidget(self.min_row_label)
        self.range_layout.addWidget(self.min_row_input)
        self.range_layout.addWidget(self.max_row_label)
        self.range_layout.addWidget(self.max_row_input)

        self.layout.addLayout(self.range_layout)
        self.layout.addWidget(self.process_btn)
        self.layout.addWidget(self.status_label)
        self.setLayout(self.layout)

        # 연결
        self.select_file_btn.clicked.connect(self.select_file)
        self.process_btn.clicked.connect(self.process_excel)

        # 파일 경로 초기화
        self.file_path = None
        self.thread = None

    def select_file(self):
        """Excel 파일 선택"""
        path, _ = QFileDialog.getOpenFileName(self, "Select Excel File")
        if path:
            self.file_path = path
            self.status_label.setText(f"Selected: {path}")

    def process_excel(self):
        """쓰레드로 Excel 작업 시작"""
        if not self.file_path:
            self.status_label.setText("Please select a source file first!")
            return

        # 사용자 입력 값 확인
        try:
            min_row = int(self.min_row_input.text())
            max_row = int(self.max_row_input.text())
        except ValueError:
            self.status_label.setText("Please enter valid integer range values.")
            return

        if min_row > max_row:
            self.status_label.setText("Min Row must be less than Max Row.")
            return

        self.status_label.setText("Starting process...")
        self.thread = ProcessThread(self.file_path, min_row, max_row)
        self.thread.progress.connect(self.update_status)
        self.thread.finished.connect(self.on_finished)
        self.thread.start()

    def update_status(self, message):
        """쓰레드에서 보낸 상태 메시지로 GUI 업데이트"""
        self.status_label.setText(message)

    def on_finished(self):
        """쓰레드 작업 완료 후 동작"""
        self.status_label.setText("Processing completed.")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    sys.exit(app.exec_())
