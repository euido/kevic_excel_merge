import sys
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel
)

#pyinstaller -w -F kevic_excel_merge.py
class ExcelFormatterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.file_path = None
        self.target_file_path = None
        self.initUI()

    def initUI(self):
        self.setWindowTitle('KEVIC')
        self.setGeometry(300, 300, 400, 300)

        self.layout = QVBoxLayout()

        self.status_label = QLabel('엑셀 파일을 선택하세요.')
        self.layout.addWidget(self.status_label)

        self.file_button = QPushButton('1.병합 엑셀 선택')
        self.file_button.clicked.connect(self.select_file)
        self.layout.addWidget(self.file_button)

        self.merge_button = QPushButton('2.병합 및 저장 (원본 파일 수정)')
        self.merge_button.clicked.connect(self.format_excel)
        self.layout.addWidget(self.merge_button)

        self.target_file_button = QPushButton('3.자재실사 양식 선택')
        self.target_file_button.clicked.connect(self.select_target_file)
        self.layout.addWidget(self.target_file_button)

        self.copy_button = QPushButton('4. 자재실사 양식 데이터 저장')
        self.copy_button.clicked.connect(self.copy_data_to_target_excel)
        self.layout.addWidget(self.copy_button)

        self.result_label = QLabel('')
        self.layout.addWidget(self.result_label)

        self.setLayout(self.layout)

    def select_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, "엑셀 파일 선택", "", "Excel Files (*.xlsx);;All Files (*)", options=options
        )
        if file_path:
            self.file_path = file_path
            self.status_label.setText(f"선택된 파일: {file_path}")

    def select_target_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, "대상 엑셀 파일 선택", "", "Excel Files (*.xlsx);;All Files (*)", options=options
        )
        if file_path:
            self.target_file_path = file_path
            self.status_label.setText(f"대상 파일 선택: {file_path}")

    def format_excel(self):
        if not self.file_path:
            self.result_label.setText("파일을 먼저 선택하세요!")
            return

        try:
            df = pd.read_excel(self.file_path, header=None)
            self.update_original_file(df)

            self.result_label.setText("원본 파일 업데이트 완료!")

        except Exception as e:
            self.result_label.setText(f"오류 발생: {e}")

    def update_original_file(self, df):
        wb = load_workbook(self.file_path)
        ws = wb.active

        current_start_row = None
        current_value = None
        merged_b_values = []  

        for index, row in df.iterrows():
            a_value = row[0]  # A열 값
            b_value = row[1]  # B열 값

            if pd.notna(a_value):
                if current_value is not None and current_start_row is not None:
                    ws.merge_cells(
                        start_row=current_start_row,
                        start_column=3,
                        end_row=index,
                        end_column=3,
                    )
                    ws.cell(row=current_start_row, column=3).value = current_value
                    ws.cell(row=current_start_row, column=3).alignment = Alignment(
                        horizontal="center", vertical="center"
                    )

                    ws.merge_cells(
                        start_row=current_start_row,
                        start_column=4,
                        end_row=index,
                        end_column=4,
                    )
                    merged_d_value = "\n".join(merged_b_values)
                    ws.cell(row=current_start_row, column=4).value = merged_d_value
                    ws.cell(row=current_start_row, column=4).alignment = Alignment(
                        horizontal="center", vertical="center", wrap_text=True
                    )

                current_start_row = index + 1
                current_value = a_value
                merged_b_values = []  

            if pd.notna(b_value):
                merged_b_values.append(str(b_value))

            ws.cell(row=index + 1, column=1).value = a_value  # A열
            ws.cell(row=index + 1, column=2).value = b_value  # B열

        if current_value is not None and current_start_row is not None:
            ws.merge_cells(
                start_row=current_start_row,
                start_column=3,
                end_row=len(df),
                end_column=3,
            )
            ws.cell(row=current_start_row, column=3).value = current_value
            ws.cell(row=current_start_row, column=3).alignment = Alignment(
                horizontal="center", vertical="center"
            )

            ws.merge_cells(
                start_row=current_start_row,
                start_column=4,
                end_row=len(df),
                end_column=4,
            )
            merged_d_value = "\n".join(merged_b_values)
            ws.cell(row=current_start_row, column=4).value = merged_d_value
            ws.cell(row=current_start_row, column=4).alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )

        wb.save(self.file_path)

    def copy_data_to_target_excel(self):
        if not self.file_path or not self.target_file_path:
            self.result_label.setText("원본 및 대상 파일을 모두 선택하세요!")
            return

        try:
            source_wb = load_workbook(self.file_path)
            source_ws = source_wb.active

            target_wb = load_workbook(self.target_file_path)
            target_ws = target_wb.active

            target_row_e = 8
            for row in range(1, source_ws.max_row + 1):
                value = source_ws.cell(row=row, column=3).value
                if value:  
                    target_ws.cell(row=target_row_e, column=5).value = value
                    target_row_e += 1

            target_row_m = 8
            for row in range(1, source_ws.max_row + 1):
                value = source_ws.cell(row=row, column=4).value
                if value:  
                    cell = target_ws.cell(row=target_row_m, column=13)
                    cell.value = value
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    target_row_m += 1

            target_wb.save(self.target_file_path)
            self.result_label.setText("데이터 복사가 완료되었습니다!")

        except Exception as e:
            self.result_label.setText(f"오류 발생: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelFormatterApp()
    window.show()
    sys.exit(app.exec_())
