import ui
import sys
from PySide6.QtWidgets import QApplication


if __name__ == '__main__':
    app = QApplication([])
    window = ui.MainWindow()
    window.show()
    sys.exit(app.exec())
