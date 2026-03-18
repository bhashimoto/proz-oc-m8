import sys

from PyQt6.QtWidgets import QApplication

from gui.main_window import MainWindow


def app():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(app())
