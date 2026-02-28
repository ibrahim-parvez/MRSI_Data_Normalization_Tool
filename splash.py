import os
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QLabel, QProgressBar, QFrame, QMessageBox
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QPixmap, QFont

# Import our new updater logic
from updater import AutoUpdater, CURRENT_VERSION, apply_update_and_restart

class StartupSplashScreen(QWidget):
    # Signal emitted when the update check is done
    startup_ready = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.setFixedSize(500, 350)
        self.setWindowFlag(Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        self.main_frame = QFrame(self)
        self.main_frame.setGeometry(0, 0, 500, 350)
        self.main_frame.setStyleSheet("""
            QFrame { background-color: #ffffff; border-radius: 15px; border: 1px solid #ddd; }
        """)
        
        layout = QVBoxLayout(self.main_frame)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setContentsMargins(40, 40, 40, 40)

        # 1. Logo
        self.logo_label = QLabel()
        self.logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.logo_label.setStyleSheet("border: none;")
        
        base_dir = os.path.dirname(os.path.abspath(__file__))
        logo_path = os.path.join(base_dir, "logo", "MRSI_excel_transformer/assets/logo.png") 
        
        if os.path.exists(logo_path):
            pixmap = QPixmap(logo_path)
            scaled_pixmap = pixmap.scaled(120, 120, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            self.logo_label.setPixmap(scaled_pixmap)
        else:
            self.logo_label.setText("MRSI")
            self.logo_label.setFont(QFont("Arial", 30, QFont.Weight.Bold))
            self.logo_label.setStyleSheet("color: #7A003C; border: none;")

        layout.addWidget(self.logo_label)
        layout.addSpacing(20)

        # 2. Title & Subtitle
        title = QLabel("Data Normalization Tool")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setFont(QFont("Arial", 18, QFont.Weight.Bold))
        title.setStyleSheet("color: #333; border: none;")
        layout.addWidget(title)

        subtitle = QLabel("McMaster Research Group for\nStable Isotopologues")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setFont(QFont("Arial", 12))
        subtitle.setStyleSheet("color: #666; border: none;")
        layout.addWidget(subtitle)
        
        layout.addSpacing(30)

        # 3. Loading Bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(8)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar { background-color: #f0f0f0; border-radius: 4px; border: none; }
            QProgressBar::chunk { background-color: #7A003C; border-radius: 4px; }
        """)
        layout.addWidget(self.progress_bar)

        # 4. Loading Text
        self.loading_text = QLabel("Checking for updates...")
        self.loading_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.loading_text.setFont(QFont("Arial", 9))
        self.loading_text.setStyleSheet("color: #888; border: none; margin-top: 5px;")
        layout.addWidget(self.loading_text)

        # --- AUTO UPDATE TRIGGER ---
        self.start_update_check()

    # --- Added missing method back ---
    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def start_update_check(self):
        self.updater_thread = AutoUpdater(mode="check")
        self.updater_thread.check_finished.connect(self.on_check_finished)
        self.updater_thread.error_occurred.connect(self.on_error)
        self.updater_thread.start()

    def on_check_finished(self, has_update, latest_version, download_url):
        if has_update:
            msg = QMessageBox(self)
            msg.setWindowTitle("Update Available")
            msg.setText(f"A newer version of the tool is available.\n\nCurrent: {CURRENT_VERSION}\nNew: {latest_version}")
            msg.setInformativeText("Would you like to update now?")
            msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            
            if msg.exec() == QMessageBox.StandardButton.Yes:
                self.start_download(download_url)
                return
        
        self.proceed_with_startup()

    def start_download(self, url):
        self.loading_text.setText("Downloading update...")
        self.progress_bar.setValue(0)
        
        self.downloader_thread = AutoUpdater(mode="download", url=url)
        self.downloader_thread.progress_updated.connect(self.progress_bar.setValue)
        self.downloader_thread.download_finished.connect(self.on_download_finished)
        self.downloader_thread.error_occurred.connect(self.on_error)
        self.downloader_thread.start()

    def on_download_finished(self, download_path):
        if download_path:
            self.loading_text.setText("Installing update and restarting...")
            apply_update_and_restart(download_path)
        else:
            self.proceed_with_startup()

    def on_error(self, error_message):
        print(error_message) 
        self.proceed_with_startup()

    def proceed_with_startup(self):
        # Tell main.py to take over and start loading the heavy modules!
        self.startup_ready.emit()