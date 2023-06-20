import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QFileDialog
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from pathlib import Path
import os
import shutil
import webbrowser


class GenerateHtmlThread(QThread):
    progress_updated = pyqtSignal(int, str)
    finished = pyqtSignal()

    def __init__(self, folder_name, data, name_column_index, phone_column_index, group_size=1000):
        super().__init__()
        self.folder_name = folder_name
        self.data = data
        self.name_column_index = name_column_index
        self.phone_column_index = phone_column_index
        self.group_size = group_size

    def run(self):
        total_groups = len(self.data) // self.group_size
        groups = [self.data[i:i + self.group_size] for i in range(0, len(self.data), self.group_size)]

        for group_index, group_data in enumerate(groups):
            html_content = "<html><body>"
            html_content+="""
                <table>
                    <tr>
                        <tr>
                         <th> Index </th>
                         <th> Name </th>
                         <th> Call </th>
                         <th> Message </th>
                         <th> WhatsApp </th>
                    </tr>
            """
            for index, row in group_data.iterrows():
                name = row[self.name_column_index - 1]
                phone_number = row[self.phone_column_index - 1]

                call_link = f"<a href='tel:{phone_number}'>Call</a>"
                message_link = f"<a href='sms:{phone_number}'>Message</a>"
                whatsapp_link = f"<a href='https://api.whatsapp.com/send?phone={phone_number}'>WhatsApp</a>"

                html_content += "<tr>"
                html_content += f"<td>{index+1}</td>"
                html_content += f"<td>{name}</td>"
                html_content += f"<td>{call_link}</td>"
                html_content += f"<td>{message_link}</td>"
                html_content += f"<td>{whatsapp_link}</td>"
                html_content += "</tr>"

            html_content += "</table>"
            html_content += "</body></html>"

            file_name = f"{self.folder_name}/contacts_group_{group_index + 1}.html"
            with open(file_name, "w", encoding="utf-8") as file:
                file.write(html_content)

            progress = int((group_index + 1) / total_groups * 100)
            self.progress_updated.emit(progress, file_name)

        self.finished.emit()


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel to HTML Converter")
        self.file_path = None
        self.name_column_index = None
        self.phone_column_index = None
        self.output_path = None
        self.generate_html_thread = None

        self.file_label = QLabel("Select Excel File:", self)
        self.file_lineedit = QLineEdit(self)
        self.file_button = QPushButton("Browse", self)

        self.name_label = QLabel("Name Column Index:", self)
        self.name_lineedit = QLineEdit(self)

        self.phone_label = QLabel("Phone Column Index:", self)
        self.phone_lineedit = QLineEdit(self)

        self.output_label = QLabel("Select Output Path:", self)
        self.output_lineedit = QLineEdit(self)
        self.output_button = QPushButton("Browse", self)

        self.convert_button = QPushButton("Convert", self)
        self.progress_label = QLabel(self)
        self.progress_label.setAlignment(Qt.AlignCenter) # type: ignore

        self.file_button.clicked.connect(self.browse_file)
        self.output_button.clicked.connect(self.browse_output_path)
        self.convert_button.clicked.connect(self.convert_excel_to_html)

        self.init_ui()

    def init_ui(self):
        self.setGeometry(100, 100, 500, 300)

        self.file_label.setGeometry(20, 20, 150, 30)
        self.file_lineedit.setGeometry(180, 20, 250, 30)
        self.file_button.setGeometry(440, 20, 50, 30)

        self.name_label.setGeometry(20, 70, 150, 30)
        self.name_lineedit.setGeometry(180, 70, 250, 30)

        self.phone_label.setGeometry(20, 120, 150, 30)
        self.phone_lineedit.setGeometry(180, 120, 250, 30)

        self.output_label.setGeometry(20, 170, 150, 30)
        self.output_lineedit.setGeometry(180, 170, 250, 30)
        self.output_button.setGeometry(440, 170, 50, 30)

        self.convert_button.setGeometry(200, 220, 100, 30)
        self.progress_label.setGeometry(20, 260, 470, 30)

    def closeEvent(self, event):
        if self.generate_html_thread is not None and self.generate_html_thread.isRunning():
            self.generate_html_thread.terminate()
            self.generate_html_thread.wait()

        event.accept()

    def browse_file(self):
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFile)
        file_dialog.setNameFilter("Excel Files (*.xlsx)")
        if file_dialog.exec_():
            file_path = file_dialog.selectedFiles()[0]
            self.file_lineedit.setText(file_path)

    def browse_output_path(self):
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.Directory)
        if file_dialog.exec_():
            output_path = file_dialog.selectedFiles()[0]
            self.output_lineedit.setText(output_path)

    def convert_excel_to_html(self):
        self.file_path = self.file_lineedit.text()
        self.name_column_index = int(self.name_lineedit.text())
        self.phone_column_index = int(self.phone_lineedit.text())
        self.output_path = self.output_lineedit.text()

        try:
            data = pd.read_excel(self.file_path)
        except Exception as e:
            print(f"Error: {e}")
            return

        folder_name = Path(self.file_path).name + '_output'
        output_folder = os.path.join(self.output_path, folder_name)
        shutil.rmtree(output_folder, ignore_errors=True)
        os.makedirs(output_folder)

        self.generate_html_thread = GenerateHtmlThread(
            output_folder, data, self.name_column_index, self.phone_column_index
        )
        self.generate_html_thread.progress_updated.connect(self.update_progress)
        self.generate_html_thread.finished.connect(self.html_generation_completed)
        self.generate_html_thread.start()

        self.convert_button.setEnabled(False)
        self.progress_label.setText("Converting...")
        self.progress_label.setStyleSheet("color: blue")

    def update_progress(self, progress, file_name):
        self.progress_label.setText(f"Generating HTML: {progress}%")

    def html_generation_completed(self):
        self.progress_label.setText("HTML generation completed.")
        self.progress_label.setStyleSheet("color: green")
        self.convert_button.setEnabled(True)

        folder_name = Path(self.file_path).name + '_output' # type: ignore
        output_folder = os.path.join(self.output_path, folder_name) # type: ignore
        webbrowser.open(f"{output_folder}/contacts_group_1.html")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
