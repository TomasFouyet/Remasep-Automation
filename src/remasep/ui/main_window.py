from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QLabel, QMainWindow, QPushButton, QVBoxLayout, QWidget


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("REMASEP Automático")
        self.resize(720, 460)

        title = QLabel("REMASEP Automático")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size: 26px; font-weight: 600;")

        status = QLabel(
            "Sprint 0\n\n"
            "La base técnica está lista.\n"
            "Las reglas y mappings reales se incorporarán después de validarlos."
        )
        status.setAlignment(Qt.AlignmentFlag.AlignCenter)

        button = QPushButton("Nuevo reporte mensual")
        button.setEnabled(False)

        layout = QVBoxLayout()
        layout.addStretch()
        layout.addWidget(title)
        layout.addSpacing(24)
        layout.addWidget(status)
        layout.addSpacing(24)
        layout.addWidget(button)
        layout.addStretch()

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)


def run_app() -> None:
    app = QApplication.instance() or QApplication([])
    window = MainWindow()
    window.show()
    app.exec()
