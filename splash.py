import os
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QLabel, QProgressBar, QFrame
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap, QFont

class StartupSplashScreen(QWidget):
    def __init__(self):
        super().__init__()
        self.setFixedSize(500, 350)
        self.setWindowFlag(Qt.WindowType.FramelessWindowHint)  # No border
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground) # Transparent corners

        # --- Main Layout (The Visible Box) ---
        self.main_frame = QFrame(self)
        self.main_frame.setGeometry(0, 0, 500, 350)
        self.main_frame.setStyleSheet("""
            QFrame {
                background-color: #ffffff;
                border-radius: 15px;
                border: 1px solid #ddd;
            }
        """)
        
        layout = QVBoxLayout(self.main_frame)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setContentsMargins(40, 40, 40, 40)

        # --- 1. Logo ---
        self.logo_label = QLabel()
        self.logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.logo_label.setStyleSheet("border: none;")
        
        # === FIX: Dynamic Path ===
        # This finds the folder where splash.py lives, then looks for 'logo/logo.png' inside it.
        base_dir = os.path.dirname(os.path.abspath(__file__))
        logo_path = os.path.join(base_dir, "logo", "/Users/iparvez/Development/Projects/MRSI/MRSI_excel_transformer/logo.png")
        
        if os.path.exists(logo_path):
            pixmap = QPixmap(logo_path)
            # Scale logo to max 120x120
            scaled_pixmap = pixmap.scaled(120, 120, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            self.logo_label.setPixmap(scaled_pixmap)
        else:
            # Fallback if logo is missing
            self.logo_label.setText("MRSI")
            self.logo_label.setFont(QFont("Arial", 30, QFont.Weight.Bold))
            self.logo_label.setStyleSheet("color: #7A003C; border: none;")

        layout.addWidget(self.logo_label)
        layout.addSpacing(20)

        # --- 2. Title ---
        title = QLabel("Data Normalization Tool")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setFont(QFont("Arial", 18, QFont.Weight.Bold))
        title.setStyleSheet("color: #333; border: none;")
        layout.addWidget(title)

        # --- 3. Subtitle ---
        subtitle = QLabel("McMaster Research Group for\nStable Isotopologues")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setFont(QFont("Arial", 12))
        subtitle.setStyleSheet("color: #666; border: none;")
        layout.addWidget(subtitle)
        
        layout.addSpacing(30)

        # --- 4. Loading Bar ---
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(8)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                background-color: #f0f0f0;
                border-radius: 4px;
                border: none;
            }
            QProgressBar::chunk {
                background-color: #7A003C; /* McMaster Maroon */
                border-radius: 4px;
            }
        """)
        layout.addWidget(self.progress_bar)

        # --- 5. Loading Text ---
        self.loading_text = QLabel("Initializing...")
        self.loading_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.loading_text.setFont(QFont("Arial", 9))
        self.loading_text.setStyleSheet("color: #888; border: none; margin-top: 5px;")
        layout.addWidget(self.loading_text)

    def update_progress(self, value):
        self.progress_bar.setValue(value)
        if value < 30:
            self.loading_text.setText("Loading modules...")
        elif value < 70:
            self.loading_text.setText("Configuring user interface...")
        elif value < 90:
            self.loading_text.setText("Checking resources...")
        else:
            self.loading_text.setText("Starting up...")