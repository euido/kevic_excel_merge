import sys
import os
import traceback
from PyQt5.QtWidgets import (
    QApplication, QVBoxLayout, QPushButton, QLabel, QFileDialog, QWidget, QLineEdit
)
from PyQt5.QtCore import QThread, pyqtSignal
from openpyxl import load_workbook


def parse_search_values(search_value):
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

    def __init__(self, file_a_path, search_values):
        super().__init__()
        self.file_a_path = file_a_path
        self.search_values = search_values  # 리스트로 숫자 범위 입력

    def run(self):
        try:
            self.progress.emit("Loading File A...")
            file_a_workbook = load_workbook(self.file_a_path)
            file_a_sheet = file_a_workbook.active

            self.progress.emit("Starting scan from B8...")

            match_count = 0  
            # B8부터 시작하여 범위 내 숫자가 있는지 확인
            for row in range(8, file_a_sheet.max_row + 1):
                # Extract values from the required columns
                b_value = file_a_sheet.cell(row=row, column=2).value  # B열
                c_value = file_a_sheet.cell(row=row, column=3).value  # C열
                d_value = file_a_sheet.cell(row=row, column=4).value  # D열
                e_value = file_a_sheet.cell(row=row, column=5).value  # E열
                i_value = file_a_sheet.cell(row=row, column=9).value  # I열
                k_value = file_a_sheet.cell(row=row, column=11).value  # K열
                m_value = file_a_sheet.cell(row=row, column=13).value  # M열

                # Log to debugging console
                if b_value is not None:
                    print(f"[DEBUG] Scanning Row {row}: B = {b_value}")
                if c_value is not None:
                    print(f"[DEBUG] Scanning Row {row}: C = {c_value}")
                if d_value is not None:
                    print(f"[DEBUG] Scanning Row {row}: D = {d_value}")
                if e_value is not None:
                    print(f"[DEBUG] Scanning Row {row}: E = {e_value}")
                if i_value is not None:
                    print(f"[DEBUG] Scanning Row {row}: I = {i_value}")
                if k_value is not None:
                    print(f"[DEBUG] Scanning Row {row}: K = {k_value}")
                if m_value is not None:
                    print(f"[DEBUG] Scanning Row {row}: M = {m_value}")

                # Perform search logic (matching user criteria)
                if any(val in self.search_values for val in [b_value, c_value, d_value, e_value, i_value, k_value, m_value]):
                    print(f"[DEBUG] Found match at Row {row}")
                    match_count += 1

                # 다중 값 확인
                if b_value in self.search_values:
                    print(f"[DEBUG] Found value at Row {row}: Value = {b_value}")
                    match_count += 1

            if match_count > 0:
                self.progress.emit(f"Found {match_count} matches.")
            else:
                self.progress.emit("No matches found.")

            self.finished.emit(match_count) 
        except Exception as e:
            error_message = f"Error: {e}"
            print("[ERROR]", error_message)
            traceback.print_exc()
            self.progress.emit(error_message)
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
        self.scan_btn = QPushButton("현품표 생성")
        
        self.status_label = QLabel("엑셀 파일을 선택하세요")

        self.layout.addWidget(self.input_label)
        self.layout.addWidget(self.search_input)
        self.layout.addWidget(self.file_a_btn)
        self.layout.addWidget(self.scan_btn)
        self.layout.addWidget(self.status_label)
        self.setLayout(self.layout)

        self.file_a_btn.clicked.connect(self.select_file_a)
        self.scan_btn.clicked.connect(self.scan_b_column)

        self.file_a_path = None
        self.thread = None

    def select_file_a(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Excel File")
        if path:
            self.file_a_path = path
            self.status_label.setText(f"Selected file: {path}")
            print(f"[DEBUG] Excel file selected: {path}")

    def scan_b_column(self):
        if not self.file_a_path:
            self.status_label.setText("Please select a file first.")
            print("[ERROR] No Excel file selected.")
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

        self.thread = ProcessThread(self.file_a_path, search_values)
        self.thread.progress.connect(self.update_status)
        self.thread.finished.connect(self.on_finished)
        self.thread.start()

    def update_status(self, message):
        print(f"[DEBUG] Status update: {message}")
        self.status_label.setText(message)

    def on_finished(self, count):
        if count > 0:
            self.status_label.setText(f"Scan complete. Found {count} matches.")
            print("[DEBUG] Matches found.")
        else:
            self.status_label.setText("Scan complete. No matches found.")
            print("[DEBUG] No matches found.")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    sys.exit(app.exec_())
