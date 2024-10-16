# Author: Leon Gajtner (optimiert)
# Datum: 15.10.2024
# PDF Magic Converter

from impl import PDFConverter, DropArea
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, 
                             QWidget, QPushButton, QProgressBar, QTextEdit, 
                             QFileDialog, QLabel, QMessageBox)
from PyQt5.QtGui import QColor, QPalette, QFont
from PyQt5.QtCore import Qt, QSettings
import sys
import os
import logging
import subprocess

class ModernButton(QPushButton):
    def __init__(self, text, color):
        super().__init__(text)
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: {color};
                border: none;
                color: white;
                padding: 10px 20px;
                text-align: center;
                text-decoration: none;
                font-size: 16px;
                margin: 4px 2px;
                border-radius: 8px;
            }}
            QPushButton:hover {{
                background-color: {QColor(color).darker(110).name()};
            }}
            QPushButton:pressed {{
                background-color: {QColor(color).darker(120).name()};
            }}
            QPushButton:disabled {{
                background-color: #CCCCCC;
                color: #666666;
            }}
        """)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF zu DOCX Konverter")
        self.setGeometry(100, 100, 800, 600)

        self.colors = {
            'background': '#F0F4F8',
            'primary': '#4A90E2',
            'secondary': '#50C878',
            'text': '#333333',
            'accent': '#FF6B6B'
        }

        self.setStyleSheet(f"""
            QMainWindow {{
                background-color: {self.colors['background']};
            }}
            QLabel, QTextEdit {{
                color: {self.colors['text']};
            }}
        """)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        title_label = QLabel("PDF zu DOCX Konverter")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet(f"""
            font-size: 24px;
            color: {self.colors['primary']};
            margin-bottom: 20px;
            font-weight: bold;
        """)
        layout.addWidget(title_label)

        self.drop_area = DropArea()
        self.drop_area.setStyleSheet(f"""
            QWidget {{
                border: 2px dashed {self.colors['primary']};
                border-radius: 12px;
                background-color: {QColor(self.colors['primary']).lighter(180).name()};
                min-height: 150px;
            }}
            QWidget:hover {{
                background-color: {QColor(self.colors['primary']).lighter(170).name()};
            }}
        """)
        self.drop_area.files_dropped.connect(self.add_files)
        layout.addWidget(self.drop_area)

        button_layout = QHBoxLayout()
        self.select_button = ModernButton("PDFs auswählen", self.colors['primary'])
        self.select_button.clicked.connect(self.select_files)
        button_layout.addWidget(self.select_button)

        self.set_output_dir_button = ModernButton("Speicherort festlegen", self.colors['accent'])
        self.set_output_dir_button.clicked.connect(self.set_output_directory)
        button_layout.addWidget(self.set_output_dir_button)

        self.convert_button = ModernButton("Konvertieren", self.colors['secondary'])
        self.convert_button.clicked.connect(self.start_conversion)
        button_layout.addWidget(self.convert_button)

        layout.addLayout(button_layout)

        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet(f"""
            QProgressBar {{
                border: 2px solid {self.colors['primary']};
                border-radius: 5px;
                text-align: center;
            }}
            QProgressBar::chunk {{
                background-color: {self.colors['secondary']};
                width: 10px;
                margin: 0.5px;
            }}
        """)
        layout.addWidget(self.progress_bar)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet(f"""
            background-color: white;
            border: 1px solid {self.colors['primary']};
            border-radius: 5px;
            padding: 5px;
        """)
        layout.addWidget(self.log_text)

        self.pdf_files = []
        self.output_directory = ""
        self.converted_files = []

        self.settings = QSettings("PDFMagicConverter", "Settings")
        self.load_settings()

        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
        self.logger = logging.getLogger(__name__)

    def load_settings(self):
        self.output_directory = self.settings.value("output_directory", "", type=str)

    def save_settings(self):
        self.settings.setValue("output_directory", self.output_directory)

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "PDFs auswählen", "", "PDF Dateien (*.pdf)")
        self.add_files(files)

    def add_files(self, files):
        new_files = [f for f in files if f not in self.pdf_files]
        self.pdf_files.extend(new_files)
        self.log_text.append(f"<span style='color: {self.colors['secondary']};'>{len(new_files)} neue PDF(s) hinzugefügt.</span>")
        self.logger.info(f"{len(new_files)} neue PDF(s) zur Konvertierung hinzugefügt.")
        self.update_drop_area_label()

    def update_drop_area_label(self):
        if self.pdf_files:
            self.drop_area.label.setText(f"<span style='font-size: 18px;'>{len(self.pdf_files)} PDF(s) ausgewählt</span>")
        else:
            self.drop_area.label.setText("<span style='font-size: 18px;'>PDF-Dateien hier hineinziehen</span>")
        self.drop_area.label.setStyleSheet(f"color: {self.colors['primary']};")

    def set_output_directory(self):
        directory = QFileDialog.getExistingDirectory(self, "Speicherort auswählen", self.output_directory)
        if directory:
            self.output_directory = directory
            self.save_settings()
            self.log_text.append(f"<span style='color: {self.colors['accent']};'>Speicherort festgelegt: {directory}</span>")

    def start_conversion(self):
        if not self.pdf_files:
            QMessageBox.warning(self, "Keine Dateien", "Bitte wählen Sie mindestens eine PDF-Datei aus.")
            return

        if not self.output_directory:
            QMessageBox.warning(self, "Kein Speicherort", "Bitte legen Sie zuerst den Speicherort fest.")
            return

        self.converted_files = []  # Reset the list of converted files
        self.converter = PDFConverter(self.pdf_files, self.output_directory)
        self.converter.update_progress.connect(self.update_progress)
        self.converter.update_log.connect(self.update_log)
        self.converter.finished.connect(self.conversion_finished)
        self.converter.start()

        self.convert_button.setEnabled(False)
        self.select_button.setEnabled(False)
        self.set_output_dir_button.setEnabled(False)

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def update_log(self, message):
        self.log_text.append(f"<span style='color: {self.colors['text']};'>{message}</span>")
        self.logger.info(message)
        if message.startswith("Erfolgreich konvertiert:"):
            pdf_file = message.split(": ")[1]
            docx_file = os.path.join(self.output_directory, os.path.splitext(os.path.basename(pdf_file))[0] + '.docx')
            self.converted_files.append(docx_file)

    def conversion_finished(self):
        self.convert_button.setEnabled(True)
        self.select_button.setEnabled(True)
        self.set_output_dir_button.setEnabled(True)
        self.pdf_files = []
        self.update_drop_area_label()
        self.log_text.append(f"<span style='color: {self.colors['secondary']};'>Konvertierung abgeschlossen.</span>")
        self.logger.info("Konvertierung abgeschlossen.")

        self.ask_to_open_files()

    def ask_to_open_files(self):
        if self.converted_files:
            reply = QMessageBox.question(self, 'Konvertierung abgeschlossen', 
                                         "Möchten Sie die konvertierten Dateien öffnen?",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            
            if reply == QMessageBox.Yes:
                self.open_converted_files()
            else:
                QMessageBox.information(self, "Konvertierung abgeschlossen", 
                                        f"Alle Dateien wurden erfolgreich konvertiert und im Verzeichnis '{self.output_directory}' gespeichert.")

    def open_converted_files(self):
        for file in self.converted_files:
            try:
                if sys.platform.startswith('darwin'):  # macOS
                    subprocess.call(('open', file))
                elif os.name == 'nt':  # Windows
                    os.startfile(file)
                elif os.name == 'posix':  # Linux
                    subprocess.call(('xdg-open', file))
            except Exception as e:
                self.logger.error(f"Fehler beim Öffnen der Datei {file}: {str(e)}")
                QMessageBox.warning(self, "Fehler", f"Konnte die Datei {file} nicht öffnen: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())