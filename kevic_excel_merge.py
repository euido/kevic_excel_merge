import sys
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel
)


class ExcelFormatterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.file_path = None
        self.initUI()

    def initUI(self):
        self.setWindowTitle('KEVIC')
        self.setGeometry(300, 300, 400, 200)

        self.layout = QVBoxLayout()

        self.status_label = QLabel('엑셀 파일을 선택하세요.')
        self.layout.addWidget(self.status_label)

        self.file_button = QPushButton('엑셀 파일 선택')
        self.file_button.clicked.connect(self.select_file)
        self.layout.addWidget(self.file_button)

        self.merge_button = QPushButton('데이터 병합 및 저장')
        self.merge_button.clicked.connect(self.format_excel)
        self.layout.addWidget(self.merge_button)

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

    def format_excel(self):
        if not self.file_path:
            self.result_label.setText("파일을 먼저 선택하세요!")
            return

        try:
            df = pd.read_excel(self.file_path, header=None)

            save_path = self.file_path.replace(".xlsx", "_Merge.xlsx")
            self.create_merged_excel(df, save_path)

            self.result_label.setText(f"저장 완료! 저장 위치: {save_path}")

        except Exception as e:
            self.result_label.setText(f"오류 발생: {e}")

    def create_merged_excel(self, df, save_path):
        wb = Workbook()
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

        wb.save(save_path)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelFormatterApp()
    window.show()
    sys.exit(app.exec_())
