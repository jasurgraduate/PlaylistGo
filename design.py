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
from PyQt6.QtGui import QColor, QIcon, QPixmap, QPainter, QLinearGradient
from PyQt6.QtCore import Qt, QSize, QTimer


class Ui_MainWindow:
    def setupUi(self, MainWindow):
        MainWindow.setWindowTitle("YouTube Song Manager")
        MainWindow.resize(1200, 800)

        self.central_widget = QWidget(MainWindow)

        # Main layout
        self.layout = QVBoxLayout(self.central_widget)
        self.layout.setContentsMargins(30, 30, 30, 30)

        # Set a cool gradient background using CSS
        gradient_css = """
            background: qlineargradient(
                x1: 0, y1: 0, x2: 1, y2: 1, 
                stop: 0 #282A36, stop: 1 #3E4D67
            );
        """
        self.central_widget.setStyleSheet(gradient_css)

        # Header (Title)
        self.header_label = QLabel("YouTube Song Manager")
        self.header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.header_label.setStyleSheet(
            "font-size: 30px; font-weight: bold; color: #FFEB3B;"
        )
        self.layout.addWidget(self.header_label)

        # Search Bar Layout
        self.search_layout = QHBoxLayout()
        self.search_label = QLabel("Search YouTube:")
        self.search_label.setStyleSheet("font-size: 18px; color: #FFEB3B;")
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Enter song or video name...")
        self.search_input.setStyleSheet(
            "font-size: 18px; padding: 12px; background-color: #2D2F3B; color: #FFEB3B; border: 2px solid #FFEB3B; border-radius: 5px;"
        )
        self.search_button = SearchButton("Search")
        self.search_layout.addWidget(self.search_label)
        self.search_layout.addWidget(self.search_input)
        self.search_layout.addWidget(self.search_button)

        # Results Layout
        self.results_layout = QVBoxLayout()

        self.results_label = QLabel("Search Results:")
        self.results_label.setStyleSheet(
            "font-size: 20px; font-weight: bold; color: #FFEB3B;"
        )
        self.results_list = QListWidget()
        self.results_list.setStyleSheet(
            "font-size: 16px; padding: 8px; background-color: #3B3F49; color: #FFEB3B; border-radius: 10px;"
        )
        self.results_list.setIconSize(
            QSize(150, 150)
        )  # Resize icons to make them bigger

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setWidget(self.results_list)

        # Add to Excel Button
        self.add_to_excel_button = QPushButton("Add Selected to Excel")
        self.add_to_excel_button.setStyleSheet(
            "font-size: 18px; padding: 12px; background-color: #007BFF; color: white; border-radius: 5px;"
        )

        # Add widgets to main layout
        self.layout.addLayout(self.search_layout)
        self.layout.addWidget(self.results_label)
        self.layout.addWidget(self.scroll_area)
        self.layout.addWidget(self.add_to_excel_button)

        MainWindow.setCentralWidget(self.central_widget)


class SearchButton(QPushButton):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.spinner_angle = 0
        self.setStyleSheet(
            "font-size: 18px; padding: 12px; background-color: #FFEB3B; color: #2D2F3B; border-radius: 5px;"
        )

    def paintEvent(self, event):
        # Paint the button with a spinning circle during the search process
        super().paintEvent(event)

        if self.isEnabled():
            return

        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(Qt.PenStyle.NoPen)

        # Draw the spinner (a circle that rotates)
        painter.setBrush(QColor(255, 255, 255))
        painter.save()
        painter.translate(self.width() // 2, self.height() // 2)
        painter.rotate(self.spinner_angle)
        painter.drawEllipse(-10, -10, 20, 20)  # Draw a small circle
        painter.restore()
