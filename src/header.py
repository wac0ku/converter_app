# Author: Leon Gajtner
# Datum: 15.10.2024
# PDF Magic Converter

# Notwendige PyQt5 Module importieren
from PyQt5.QtCore import pyqtSignal, QThread
from PyQt5.QtWidgets import QWidget, QLabel, QVBoxLayout
from PyQt5.QtGui import QDropEvent, QDragEnterEvent

# Interfaces für PDF-Konverter und DropArea
class PDFConverterInterface(QThread):
    update_progress = pyqtSignal(int)
    update_log = pyqtSignal(str)
    
    def __init__(self, pdf_files):
        super().__init__()
        self.pdf_files = pdf_files

    def run(self):
        raise NotImplementedError("Methode 'run' muss in der abgeleiteten Klasse implementiert werden.")


class DropAreaInterface(QWidget):
    files_dropped = pyqtSignal(list)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event: QDragEnterEvent):
        raise NotImplementedError("Methode 'dragEnterEvent' muss in der abgeleiteten Klasse implementiert werden.")

    def dragMoveEvent(self, event):
        raise NotImplementedError("Methode 'dragMoveEvent' muss in der abgeleiteten Klasse implementiert werden.")

    def dropEvent(self, event: QDropEvent):
        raise NotImplementedError("Methode 'dropEvent' muss in der abgeleiteten Klasse implementiert werden.")
