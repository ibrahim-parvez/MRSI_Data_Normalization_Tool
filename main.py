import sys
from PyQt6.QtWidgets import QApplication
from gui import DataToolApp

def main():
    app = QApplication(sys.argv)
    window = DataToolApp()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
