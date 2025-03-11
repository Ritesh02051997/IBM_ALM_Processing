import sys
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QLineEdit, QMainWindow, QTextEdit
from PyQt6.QtGui import QAction, QPalette, QPixmap, QBrush
from PyQt6.QtCore import Qt, QUrl
from PyQt6.QtGui import QDesktopServices
from os import path
import SWIT2CALIT_main

class ALMAutomationApp(QMainWindow):
    def __init__(self):
        super().__init__()

        # Set up the main window
        self.setWindowTitle('SW-IT 2 CAL-IT Automation')
        self.setGeometry(350, 250, 1200, 700)

        # Set Background Image
        self.set_background_image()

        # Main widget
        self.main_widget = QWidget(self)
        self.setCentralWidget(self.main_widget)

        # Layouts
        self.main_layout = QVBoxLayout()
        self.button_layout = QVBoxLayout()
        self.log_layout = QVBoxLayout()

        # Heading label
        self.heading_label = QLabel("SW-IT 2 CAL-IT Automation Tool")
        self.heading_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.heading_label.setStyleSheet("font-weight: bold; font-size: 32px; color: purple;")

        # Menu bar
        self.menu_bar = self.menuBar().setStyleSheet("QMenuBar {background-color: grey;}")
        self.create_menu()

        # CSV and Mapping file labels and buttons
        self.csv_label = QLineEdit(self)
        self.csv_label.setReadOnly(True)
        self.csv_label.setPlaceholderText('No file selected...')
        self.csv_label.setStyleSheet("font-size: 15px; font-weight: bold;")
        self.csv_button = QPushButton('Input CSV', self)
        self.csv_button.clicked.connect(self.upload_csv)
        self.csv_button.setStyleSheet("font-size: 15px; font-weight: bold;")

        self.mapping_label = QLineEdit(self)
        self.mapping_label.setReadOnly(True)
        self.mapping_label.setPlaceholderText('No file selected...')
        self.mapping_label.setStyleSheet("font-size: 15px; font-weight: bold;")
        self.mapping_button = QPushButton('Input Config', self)
        self.mapping_button.clicked.connect(self.upload_mapping)
        self.mapping_button.setStyleSheet("font-size: 15px; font-weight: bold;")

        # Output directory label and buttons
        self.output_label = QLineEdit(self)
        self.output_label.setReadOnly(True)
        self.output_label.setPlaceholderText('No directory selected...')
        self.output_label.setStyleSheet("font-size: 15px; font-weight: bold;")
        self.output_button = QPushButton('Select Output Directory', self)
        self.output_button.clicked.connect(self.select_output_directory)
        self.output_button.setStyleSheet("font-size: 15px; font-weight: bold;")
        
        self.open_output_button = QPushButton('Open Output Folder', self)
        self.open_output_button.clicked.connect(self.open_output_folder)
        self.open_output_button.setStyleSheet("font-size: 15px; font-weight: bold;")

        # Generate CSV Button
        self.generate_button = QPushButton('Generate CSV', self)
        self.generate_button.clicked.connect(self.generate_csv)
        self.generate_button.setStyleSheet("font-size: 18px; font-weight: bold; color: orange;")

        # Log display
        self.log_display = QTextEdit(self)
        self.log_display.setReadOnly(True)
        self.log_display.setStyleSheet("font-size: 15px; font-weight: bold;")

        # Reset Button
        self.reset_button = QPushButton('Reset', self)
        self.reset_button.clicked.connect(self.reset)
        self.reset_button.setStyleSheet("font-size: 15px; font-weight: bold;")

        # Adding widgets to layout
        self.button_layout.addWidget(self.csv_label)
        self.button_layout.addWidget(self.csv_button)
        self.button_layout.addWidget(self.mapping_label)
        self.button_layout.addWidget(self.mapping_button)
        self.button_layout.addWidget(self.output_label)
        self.button_layout.addWidget(self.output_button)
        self.button_layout.addWidget(self.generate_button)
        self.button_layout.addWidget(self.open_output_button)

        self.log_layout.addWidget(self.log_display)
        self.log_layout.addWidget(self.reset_button)

        self.main_layout.addWidget(self.heading_label)
        self.main_layout.addLayout(self.button_layout)
        self.main_layout.addLayout(self.log_layout)

        self.main_widget.setLayout(self.main_layout)

    def set_background_image(self):
        # Get the path to the image file
        if getattr(sys, 'frozen', False):
            # If running as a bundled exe, get path from sys._MEIPASS
            image_path = path.join(sys._MEIPASS, 'test1.jpg')
        else:
            # If running in development, use the normal path
            image_path = 'test1.jpg'
        # Set background image using QPalette
        palette = QPalette()
        pixmap = QPixmap(image_path)  # Provide the path to your image file
        pixmap = pixmap.scaled(self.size(), aspectRatioMode=Qt.AspectRatioMode.IgnoreAspectRatio)
        brush = QBrush(pixmap)  # Create a QBrush with the QPixmap
        palette.setBrush(QPalette.ColorRole.Window, brush)  # Use ColorRole.Window for the main window background
        self.setPalette(palette)
        self.setAutoFillBackground(True)

    def create_menu(self):
        # Create the 'About' and 'Feedback' menu options
        about_action = QAction("About", self)
        about_action.triggered.connect(self.show_about)
        feedback_action = QAction("Feedback", self)
        feedback_action.triggered.connect(self.send_feedback)

        # Add the menu to the menu bar
        menu = self.menuBar().addMenu("Help")
        menu.addAction(about_action)
        menu.addAction(feedback_action)  # Add Feedback action to Help menu

    def show_about(self):
        # Show developer details in a popup
        about_message = "You are using SW-IT 2 CAL-IT Automation Tool.\nVersion:1.0.0\nDeveloper: Ritesh Sinha\nNTID: IIH3KOR"
        self.show_log_message(about_message)

    def send_feedback(self):
        # Open default email client with a pre-configured email
        feedback_url = QUrl("mailto:ritesh.sinha2@in.bosch.com?subject=SW-IT 2 CAL-IT Automation Feedback&body=More than happy to hear from you. Please share your feedbacks.")
        QDesktopServices.openUrl(feedback_url)

    def upload_csv(self):
        # Open file dialog to select a .csv file
        file, _ = QFileDialog.getOpenFileName(self, "Open CSV File", "", "CSV Files (*.csv);")
        if file:
            self.csv_label.setText(file)
            self.output_label.setText(path.dirname(file))
            self.show_log_message(f"CSV file uploaded: {path.basename(file)}")

    def upload_mapping(self):
        # Open file dialog to select a mapping .xlsx file
        file, _ = QFileDialog.getOpenFileName(self, "Open Mapping File", "", "Excel Files (*.xlsx)")
        if file:
            self.mapping_label.setText(file)
            self.show_log_message(f"Mapping file uploaded: {path.basename(file)}")

    def select_output_directory(self):
        # Open file dialog to select an output directory
        directory = QFileDialog.getExistingDirectory(self, "Select Output Directory")
        if directory:
            self.output_label.setText(directory)
            self.show_log_message(f"Alright! Got my destiny: {directory}")

    def open_output_folder(self):
        # Open the selected output folder in the file explorer
        output_directory = self.output_label.text()
        if path.exists(output_directory):
            QDesktopServices.openUrl(QUrl.fromLocalFile(output_directory))
            self.show_log_message(f"Opening Output Folder: {output_directory}")
        else:
            self.show_log_message(f"Error: Output directory does not exist.", "red")

    def generate_csv(self):
        try:
            # Placeholder: Read the uploaded CSV and Mapping (Excel) files
            csv_file = self.csv_label.text()
            mapping_file = self.mapping_label.text()
            output_dir = self.output_label.text()

            if not csv_file or not mapping_file or not output_dir:
                raise ValueError("All fields (CSV, Mapping, and Output Directory) must be filled before generating output.")

            self.show_log_message(f"Good things take time.", "green")
            self.show_log_message(f"Processing input file.", "green")
            alm_df = SWIT2CALIT_main.filter_excel(csv_file).exceltodf()
            self.show_log_message(f"I am mid way through", "green")
            alm_df_update = SWIT2CALIT_main.mapping_data(mapping_file, alm_df).process_main()
            self.show_log_message(f"Almost there!", "green")
            output_file = SWIT2CALIT_main.output_excel(mapping_file, alm_df_update, output_dir).generate_main()
            self.show_log_message(f"All done.", "green")
            self.show_log_message(f"Congratulations, output generated: {output_file}", "green")

        except Exception as e:
            self.show_log_message(f"Error: {str(e)}", "red")

    def reset(self):
        # Reset all fields and clear the log
        self.csv_label.clear()
        self.mapping_label.clear()
        self.output_label.clear()
        self.log_display.clear()

    def show_log_message(self, message, color="black"):
        # Display log message with optional color
        self.log_display.setTextColor(Qt.GlobalColor.green if color == "green" else Qt.GlobalColor.red if color == "red" else Qt.GlobalColor.darkGray)
        self.log_display.append(message)


# Main entry point for the application
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ALMAutomationApp()
    window.show()
    sys.exit(app.exec())
