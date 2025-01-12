from PyQt6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QListWidget,
    QScrollArea,
)
from PyQt6.QtCore import Qt, QSize


class Ui_MainWindow:
    def setupUi(self, MainWindow):
        MainWindow.setWindowTitle("YouTube Song Manager")
        MainWindow.resize(1200, 800)

        self.central_widget = QWidget(MainWindow)

        self.layout = QVBoxLayout(self.central_widget)
        self.layout.setContentsMargins(30, 30, 30, 30)

        gradient_css = """
            background: qlineargradient(
                x1: 0, y1: 0, x2: 1, y2: 1, 
                stop: 0 #1e3c72, stop: 1 #2a5298
            );
            border-radius: 15px;
        """
        self.central_widget.setStyleSheet(gradient_css)

        self.header_label = QLabel("YouTube Song Manager")
        self.header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.header_label.setStyleSheet(
            "font-size: 36px; font-weight: bold; color: #ffffff; margin-bottom: 20px;"
        )
        self.layout.addWidget(self.header_label)

        self.search_layout = QHBoxLayout()
        self.search_label = QLabel("Search YouTube:")
        self.search_label.setStyleSheet("font-size: 20px; color: #ffffff;")
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Enter song or video name...")
        self.search_input.setStyleSheet(
            "font-size: 18px; padding: 12px; background-color: #ffffff; color: #000000; border: 2px solid #ffffff; border-radius: 5px;"
        )
        self.search_button = QPushButton("Search")
        self.search_button.setStyleSheet(
            "font-size: 18px; padding: 12px; background-color: #ff6347; color: #ffffff; border-radius: 5px;"
        )
        self.search_layout.addWidget(self.search_label)
        self.search_layout.addWidget(self.search_input)
        self.search_layout.addWidget(self.search_button)

        self.results_label = QLabel("Search Results:")
        self.results_label.setStyleSheet(
            "font-size: 22px; font-weight: bold; color: #ffffff; margin-top: 20px;"
        )
        self.results_list = QListWidget()
        self.results_list.setStyleSheet(
            """
            QListWidget {
                font-size: 16px; padding: 10px; background-color: #ffffff; color: #000000; border-radius: 10px;
            }
            QListWidget::item {
                margin: 5px;
    padding: 2px;
    border-radius: 10px;
    background-color: #f0f0f0;
            }
            QListWidget::item:hover {
                background-color: rgba(255, 99, 71, 0.05);
                border: 0.1px solid #ff6347;
                border-radius: 10px;
            }
            QListWidget::item:selected {
                background-color: rgba(0, 41, 106, 0.532);
    color: white;
    border: 0.1px solid #ff32e0;
            }
            """
        )
        self.results_list.setIconSize(QSize(100, 100))

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setWidget(self.results_list)

        self.add_to_excel_button = QPushButton("Add Selected to Excel")
        self.add_to_excel_button.setStyleSheet(
            "font-size: 18px; padding: 12px; background-color: #32cd32; color: #ffffff; border-radius: 5px; margin-top: 20px;"
        )

        self.layout.addLayout(self.search_layout)
        self.layout.addWidget(self.results_label)
        self.layout.addWidget(self.scroll_area)
        self.layout.addWidget(self.add_to_excel_button)

        MainWindow.setCentralWidget(self.central_widget)
