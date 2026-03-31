import sys

from PyQt6.QtWidgets import QApplication

from gui.main_window import MainWindow


def app():
    try:
        app = QApplication(sys.argv)
        window = MainWindow()
        window.show()
        return app.exec()
    except Exception as e:
        with open("error.txt", "w") as f:
            f.write(e)
        sys.exit()


if __name__ == "__main__":
    sys.exit(app())
