import sys
from PySide6 import QtWidgets
from PySide6.QtCore import Qt

def main():
    app = QtWidgets.QApplication(sys.argv)

    window = QtWidgets.QWidget()
    window.setWindowTitle("Resolve Hello World")

    layout = QtWidgets.QVBoxLayout(window)

    label = QtWidgets.QLabel("Hello World")
    label.setAlignment(Qt.AlignCenter)

    layout.addWidget(label)

    window.resize(300, 120)
    window.show()

    sys.exit(app.exec())

if __name__ == "__main__":
    main()
