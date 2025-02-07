import sys
import os
from PyQt6.QtCore import QThread, pyqtSignal
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (QApplication, QWidget, QPushButton, QLineEdit, QLabel,
                             QGridLayout, QFileDialog, QComboBox, QMessageBox)
from PyQt6.QtGui import QPixmap, QFont, QIcon
from redact import *
from prompting import *
from time import sleep
from global_signals import global_signals
from write_report_mcp import clean_up

# Import write_report modules (MCP and DATA)
import write_report_mcp as mcp_write_report
import write_report_data as data_write_report

##For debugging only:
#print('YOU NEED TO COMMENT THIS OUT BEFORE EXE')
#script_directory = os.path.dirname(os.path.abspath(__file__))
#os.chdir(script_directory)

# Function to get the correct path for accessing resources in a bundled app
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Define paths for resources
logo_path_abs = "resources/ormittalentV3.png"
icon_path_abs = "resources/assessmentReport.ico"

logo_path = resource_path(logo_path_abs)
icon_path = resource_path(icon_path_abs)

programs = ['MCP', 'DATA']
genders = ['M', 'F']

class ProcessingThread(QThread):
    processing_completed = pyqtSignal(str)  # Signal to emit when processing is done

    def __init__(self, GUI_data):
        super().__init__()
        self.GUI_data = GUI_data

    def run(self):
        # Redact and store files
        redact_folder(self.GUI_data)
        print('Redaction completed')

        # Send prompts to OpenAI
        output_path = send_prompts(self.GUI_data)
        print('Drafting completed')

        # Convert JSON to report
        clean_data = clean_up(output_path)
        selected_program = self.GUI_data["Traineeship"]
        if selected_program == 'MCP':
            updated_doc = mcp_write_report.update_document(clean_data, self.GUI_data["Applicant Name"], self.GUI_data["Assessor Name"], self.GUI_data["Gender"], self.GUI_data["Traineeship"])
        elif selected_program == 'DATA':
            updated_doc = data_write_report.update_document(clean_data, self.GUI_data["Applicant Name"], self.GUI_data["Assessor Name"], self.GUI_data["Gender"], self.GUI_data["Traineeship"])
        else:
            updated_doc = mcp_write_report.update_document(clean_data, self.GUI_data["Applicant Name"], self.GUI_data["Assessor Name"], self.GUI_data["Gender"], self.GUI_data["Traineeship"]) # Default to MCP if program is not recognized
        print('Writing completed')

        # Emit the path of the generated document
        self.processing_completed.emit(updated_doc)

class MainWindow(QWidget):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setWindowTitle("ORMIT - Draft Assessment Report v1.0")
        self.setWindowIcon(QIcon(icon_path))
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowType.WindowStaysOnTopHint)
        # self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)
        self.setFixedWidth(1000)
        self.setStyleSheet("background-color: white; color: black;")
        bold_font = QFont()
        bold_font.setBold(True)

        layout = QGridLayout()
        self.setLayout(layout)

        # Initialize the message box once here
        self.msg_box = QMessageBox(self)
        self.msg_box.setWindowTitle("Processing")
        self.msg_box.setStandardButtons(QMessageBox.StandardButton.NoButton)
        self.msg_box.setWindowFlags(self.msg_box.windowFlags() | Qt.WindowType.WindowMinimizeButtonHint)
        self.msg_box.setStandardButtons(QMessageBox.StandardButton.Close)  # Add Close button
        self.msg_box.button(QMessageBox.StandardButton.Close).clicked.connect(self.close_application)

        global_signals.update_message.connect(self.refresh_message_box)

        # Load the logo
        pixmap = QPixmap(logo_path)
        pixmap_label = QLabel()
        pixmap_label.setScaledContents(True)
        resize_fac = 3
        scaled_pixmap = pixmap.scaled(
            round(pixmap.width() / resize_fac),
            round(pixmap.height() / resize_fac),
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation
        )
        pixmap_label.setPixmap(scaled_pixmap)
        layout.addWidget(pixmap_label, 0, 0, 1, 2)

        # OpenAI Key input
        self.key_label = QLabel('Gemini Key:')
        self.key_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.key_label, 1, 0)

        self.openai_key_input = QLineEdit(placeholderText='Enter Gemini Key: ***************')
        layout.addWidget(self.openai_key_input, 1, 1, 1, 2)

        # Applicant information
        self.key_label = QLabel('Applicant:')
        self.key_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.key_label, 2, 0)

        self.applicant_name_input = QLineEdit(placeholderText='Applicant Full Name')
        layout.addWidget(self.applicant_name_input, 2, 1, 1, 2)

        # Assessor information
        self.key_label = QLabel('Assessor:')
        self.key_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.key_label, 3, 0)

        self.assessor_name_input = QLineEdit(placeholderText='Assessor Full Name')
        layout.addWidget(self.assessor_name_input, 3, 1, 1, 2)

        # Select Gender
        cat_label = QLabel('Gender:')
        layout.addWidget(cat_label, 5, 0)

        self.combo_title = QComboBox(self)
        for i in genders:
            self.combo_title.addItem(i)
        self.combo_title.currentIndexChanged.connect(lambda: self.selectionchange_traineeship(self.combo_title))
        self.combo_title.setToolTip('Select a gender')
        layout.addWidget(self.combo_title, 4, 0)

        # Select Traineeship
        cat_label = QLabel('Traineeship:')
        layout.addWidget(cat_label, 5, 0)

        self.combo_title2 = QComboBox(self)
        for i in programs:
            self.combo_title2.addItem(i)
        self.combo_title2.currentIndexChanged.connect(lambda: self.selectionchange_traineeship(self.combo_title2))
        self.combo_title2.setToolTip('Select a traineeship')
        layout.addWidget(self.combo_title2, 5, 0)

        # Document labels
        self.file_label1 = QLabel("No file selected", self)
        self.file_label2 = QLabel("No file selected", self)
        self.file_label3 = QLabel("No file selected", self)

        # File selection buttons
        self.selected_files = {}  # Dictionary to store the full file paths
        self.file_browser_btn = QPushButton('PAPI Feedback')
        self.file_browser_btn.clicked.connect(lambda: self.open_file_dialog(1))
        layout.addWidget(self.file_browser_btn, 6, 0)
        layout.addWidget(self.file_label1, 6, 1, 1, 3)

        self.file_browser_btn2 = QPushButton('Cog. Test')
        self.file_browser_btn2.clicked.connect(lambda: self.open_file_dialog(2))
        layout.addWidget(self.file_browser_btn2, 7, 0)
        layout.addWidget(self.file_label2, 7, 1, 1, 3)

        self.file_browser_btn3 = QPushButton('Assessment Notes')
        self.file_browser_btn3.clicked.connect(lambda: self.open_file_dialog(3))
        layout.addWidget(self.file_browser_btn3, 8, 0)
        layout.addWidget(self.file_label3, 8, 1, 1, 3)

        # Submit button
        self.submitbtn = QPushButton('Submit')
        self.submitbtn.setFixedWidth(90)
        self.submitbtn.hide()  # Initially hidden
        layout.addWidget(self.submitbtn, 9, 4)
        self.submitbtn.clicked.connect(self.handle_submit)  # Connect submit button to handler

        # Counter for selected files
        self.selected_files_count = 0

    def refresh_message_box(self, message):
        self.msg_box.setText(message)
        self.msg_box.show()
    def close_application(self):
        # This will close the application when the messagebox is closed manually
        QApplication.quit()

    def open_file_dialog(self, file_index):
        dialog = QFileDialog(self)
        dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
        dialog.setNameFilter("PDF Files (*.pdf);;All Files (*)")
        dialog.setViewMode(QFileDialog.ViewMode.List)

        if dialog.exec():
            filenames = dialog.selectedFiles()
            if filenames:
                selected_file = str(filenames[0])
                if file_index == 1:
                    self.file_label1.setText(os.path.basename(selected_file))
                    self.selected_files["PAPI Feedback"] = selected_file  # Store the full path
                elif file_index == 2:
                    self.file_label2.setText(os.path.basename(selected_file))
                    self.selected_files["Cog. Test"] = selected_file
                elif file_index == 3:
                    self.file_label3.setText(os.path.basename(selected_file))
                    self.selected_files["Assessment Notes"] = selected_file

                self.selected_files_count += 1  # Increment counter for selected files

                # Show submit button only if all three files have been selected
                if self.selected_files_count == 3:
                    self.submitbtn.show()

    def selectionchange_traineeship(self, i):
        program_final = i.currentText()
        print("Selection changed: ", i.currentText())

    def handle_submit(self):
        # Gather all the data into a dictionary
        GUI_data = {
            "Gemini Key": self.openai_key_input.text(),  # Corrected key name to "Gemini Key"
            "Applicant Name": self.applicant_name_input.text(),
            "Assessor Name": self.assessor_name_input.text(),
            "Gender": self.combo_title.currentText(),
            "Traineeship": self.combo_title2.currentText(),
            "Files": {
                "PAPI Feedback": self.selected_files.get("PAPI Feedback", ""),
                "Cog. Test": self.selected_files.get("Cog. Test", ""),
                "Assessment Notes": self.selected_files.get("Assessment Notes", "")
            }
        }

        # Start the processing thread
        self.processing_thread = ProcessingThread(GUI_data)
        self.processing_thread.processing_completed.connect(self.on_processing_completed)
        self.processing_thread.start()

    def on_processing_completed(self, updated_doc):
        self.msg_box.close()

        os.startfile(updated_doc)

        self.close()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())  # Event loop terminates when MainWindow closes
