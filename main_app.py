
import sys
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QPushButton,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QFrame,
)
from PySide6.QtGui import QFont, Qt, QPixmap


from tools.file_preprocess import run_file_preprocessing_workflow
from tools.template_comparator import run_template_workflow
from tools.field_report import run_field_report_workflow


class ToolsApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("LEDA Group - Implementation Tools")
        self.setFixedSize(480, 380)

        # ---------- Fonts ----------
        #header_font = QFont("Arial", 14, QFont.Bold)
        button_font = QFont("Arial", 12, QFont.Bold)

        # ---------- Header with Logo Placeholder ----------
        header_frame = QFrame()
        header_frame.setStyleSheet("background-color: #1F2933;")

        # HBox → stretches
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.setSpacing(0)
        header_frame.setLayout(header_layout)

        # VBox inside → moves logo to the top
        logo_container = QVBoxLayout()
        logo_container.setContentsMargins(0, 2, 0, 0)  # margin from the top
        logo_container.setSpacing(0)

        logo_label = QLabel()
        pixmap = QPixmap("assets/logo.png")
        pixmap = pixmap.scaled(120, 60, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        logo_label.setPixmap(pixmap)

        # Smaller fixed area to cut height in half
        logo_label.setFixedSize(132, 66)
        logo_label.setStyleSheet("background-color: none;")

        # Add to vertical container (pins top)
        logo_container.addWidget(logo_label, alignment=Qt.AlignTop | Qt.AlignHCenter)

        # Add stretch under logo to keep it at top
        logo_container.addStretch()

        # Add the container to the header layout, centered horizontally
        header_layout.addStretch()
        header_layout.addLayout(logo_container)
        header_layout.addStretch()

        # ---------- Buttons ----------
        self.file_preprocessing_button = QPushButton("File Preprocessing")
        self.field_report_button = QPushButton("Field Report")
        self.tool3_button = QPushButton("Tool 3")

        for btn in (self.file_preprocessing_button, self.field_report_button, self.tool3_button):
            btn.setFont(button_font)
            btn.setFixedHeight(70)
            btn.setStyleSheet(
                """
                QPushButton {
                    background-color: #27AE60;
                    color: white;
                    border-radius: 6px;
                }
                QPushButton:hover {
                    background-color: #1E8449;
                }
                QPushButton:pressed {
                    background-color: #145A32;
                }
                """
            )

        # Wire buttons
        self.file_preprocessing_button.clicked.connect(
            lambda: run_file_preprocessing_workflow(self)
        )
        self.field_report_button.clicked.connect(
            lambda: run_field_report_workflow(self)
        )
        self.tool3_button.clicked.connect(
            lambda: run_template_workflow(self)
        )

        # ---------- Central Layout ----------
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)

        main_layout.addWidget(header_frame)
        main_layout.addSpacing(15)
        main_layout.addWidget(self.file_preprocessing_button)
        main_layout.addSpacing(2)
        main_layout.addWidget(self.field_report_button)
        main_layout.addSpacing(2)
        main_layout.addWidget(self.tool3_button)
        main_layout.addStretch()

        self.setCentralWidget(main_widget)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ToolsApp()
    window.show()
    sys.exit(app.exec())
