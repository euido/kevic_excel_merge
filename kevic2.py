import sys
import os
import traceback
from PyQt5.QtWidgets import (
    QApplication, QVBoxLayout, QPushButton, QLabel, QFileDialog, QWidget, QLineEdit
)
from PyQt5.QtCore import QThread, pyqtSignal
from openpyxl import load_workbook

def parse_search_values(search_value):
    """Parsing the search range string into a list of integers."""
    values = set()
    for part in search_value.split(","):
        part = part.strip()
        if "-" in part:
            try:
                start, end = map(int, part.split("-"))
                values.update(range(start, end + 1))
            except ValueError:
                pass
        else:
            try:
                values.add(int(part))
            except ValueError:
                pass
    return list(values)

class ProcessThread(QThread):
    progress = pyqtSignal(str)
    finished = pyqtSignal(int)

    def __init__(self, file_a_path, search_values, file_b_path):
        super().__init__()
        self.file_a_path = file_a_path
        self.search_values = search_values
        self.file_b_path = file_b_path

    def run(self):
        try:
            self.progress.emit("Loading File A and File B...")
            file_a_workbook = load_workbook(self.file_a_path)
            file_a_sheet = file_a_workbook.active

            file_b_workbook = load_workbook(self.file_b_path)
            base_sheet = file_b_workbook.active

            self.progress.emit("Starting scan from B8...")
            match_count = 0

            for row in range(8, file_a_sheet.max_row + 1):
                b_value = file_a_sheet.cell(row=row, column=2).value  # B열

                if b_value in self.search_values:
                    self.progress.emit(f"Match found at row {row}")
                    match_count += 1

                    new_sheet_name = f"{b_value}"

                    if new_sheet_name in file_b_workbook.sheetnames:
                        self.progress.emit(f"Sheet {new_sheet_name} already exists. Skipping.")
                        continue

                    new_sheet = file_b_workbook.copy_worksheet(base_sheet)
                    new_sheet.title = new_sheet_name
                

            file_b_workbook.save(self.file_b_path)
            self.progress.emit("All matches processed and saved.")
            self.finished.emit(match_count)

        except Exception as e:
            self.progress.emit(f"Error: {str(e)}")
            self.finished.emit(0)

class MainApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("KEVIC 현품표 Ver1.0.0")
        self.resize(500, 300)

        self.layout = QVBoxLayout()

        self.input_label = QLabel("(2,3,4 or 2-5):")
        self.search_input = QLineEdit()
        self.file_a_btn = QPushButton("현황표를 선택하세요")
        self.file_b_btn = QPushButton("현품표 서식을 선택하세요 ")
        self.scan_btn = QPushButton("현품표 생성")
        self.status_label = QLabel("엑셀 파일을 선택하세요")

        self.layout.addWidget(self.input_label)
        self.layout.addWidget(self.search_input)
        self.layout.addWidget(self.file_a_btn)
        self.layout.addWidget(self.file_b_btn)
        self.layout.addWidget(self.scan_btn)
        self.layout.addWidget(self.status_label)
        self.setLayout(self.layout)

        self.file_a_btn.clicked.connect(self.select_file_a)
        self.file_b_btn.clicked.connect(self.select_file_b)
        self.scan_btn.clicked.connect(self.scan_b_column)

        self.file_a_path = None
        self.file_b_path = None
        self.thread = None

    def select_file_a(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select File A")
        if path:
            self.file_a_path = path
            self.status_label.setText(f"Selected File A: {path}")
            print(f"[DEBUG] Selected File A: {path}")

    def select_file_b(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select B Excel File")
        if path:
            self.file_b_path = path
            self.status_label.setText(f"Selected B File: {path}")
            print(f"[DEBUG] Selected B File: {path}")

    def scan_b_column(self):
        if not self.file_a_path or not self.file_b_path:
            self.status_label.setText("Please select both Excel files first.")
            print("[ERROR] Missing required files.")
            return

        search_value = self.search_input.text()
        if not search_value.strip():
            self.status_label.setText("Please enter search values.")
            print("[ERROR] No search values entered.")
            return

        try:
            search_values = parse_search_values(search_value)
        except Exception as e:
            self.status_label.setText("Error parsing input.")
            print("[ERROR] Failed to parse input.", e)
            return

        if self.thread and self.thread.isRunning():
            self.thread.quit()
            self.thread.wait()

        self.thread = ProcessThread(self.file_a_path, search_values, self.file_b_path)
        self.thread.progress.connect(self.update_status)
        self.thread.finished.connect(self.on_finished)
        self.thread.start()

    def update_status(self, message):
        print(f"[DEBUG] Status: {message}")
        self.status_label.setText(message)

    def on_finished(self, count):
        if count > 0:
            self.status_label.setText(f"Created {count} matching sheets successfully.")
        else:
            self.status_label.setText("No matches or failed to generate.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    sys.exit(app.exec_())
