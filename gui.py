import sys
import os
import subprocess
import time
import pandas as pd
import xlwings as xw
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QFileDialog, QTabWidget, QTextEdit, QCheckBox,
    QLineEdit, QComboBox, QGroupBox, QMessageBox, QMenu, QProgressBar, QFrame,
    QSizePolicy, QSpacerItem, QGridLayout, QTabBar, QDialog
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer, QPoint
from PyQt6.QtGui import QFont, QIcon
# --- Import Tab UIs ---
from carbonate_tab import CarbonateTab
from water_tab import WaterTab
# --- NEW/UPDATED IMPORT ---
import settings
import hashlib

# ---- Import your step modules (Keep these imports for the WorkerThread) ----
from steps.carbonate.step1_data import step1_data_carbonate
from steps.carbonate.step2_tosort import step2_tosort_carbonate
from steps.carbonate.step3_last6 import step3_last6_carbonate
from steps.carbonate.step4_pre_group import step4_pre_group_carbonate
from steps.carbonate.step5_group import step5_group_carbonate
from steps.carbonate.step6_normalization import step6_normalization_carbonate
from steps.carbonate.step7_report import step7_report_carbonate
from steps.water.step1_data import step1_data_water
from steps.water.step2_tosort import step2_tosort_water
from steps.water.step3_last6 import step3_last6_water
from steps.water.step4_pre_group import step4_pre_group_water
from steps.water.step5_group import step5_group_water
from steps.water.step6_normalization import step6_normalization_water
from steps.water.step7_report import step7_report_water

# ---------------- Utility: XLS → XLSX ----------------
def convert_xls_to_xlsx(file_path):
    """
    Converts .xls to .xlsx and returns the new file path.
    Raises an exception if conversion fails.
    """
    new_path = os.path.splitext(file_path)[0] + ".xlsx"
    try:
        df = pd.read_excel(file_path, engine="xlrd")
        with pd.ExcelWriter(new_path, engine="openpyxl") as writer:
            default_sheet_name = "Default_Gas_Bench.wke"
            df.to_excel(writer, index=False, sheet_name=default_sheet_name)
        return new_path
    except Exception as e:
        raise Exception(f"XLS to XLSX conversion failed: {e}")

def refresh_excel(file_path):
    app = xw.App(visible=False)
    try:
        wb = app.books.open(os.path.abspath(file_path))
        wb.app.calculate()
        time.sleep(1)
        wb.save()
        wb.close()
    finally:
        app.quit()

# ---------------- Worker Thread ----------------
class WorkerThread(QThread):
    log = pyqtSignal(str, str)
    progress = pyqtSignal(int, int, str)
    stopped_early = pyqtSignal() # New signal for user stop
    
    def __init__(self, file_path, steps, sheet_name, filter_option, tab_type="carbonate"):
        super().__init__()
        self.file_path = file_path
        self.steps = steps
        self.sheet_name = sheet_name
        self.filter_option = filter_option
        self.tab_type = tab_type
        self._is_running = True # Control flag

    def stop(self):
        """Safely stops the thread *before* the next step starts."""
        self._is_running = False
        
    def run(self):
        try:
            # 1. Define Step Functions
            #step1_func = step1_data_carbonate if self.tab_type == "Carbonate" else step1_data_water
            step2_carbonate = lambda: step2_tosort_carbonate(self.file_path, self.filter_option)
            step2_water = lambda: step2_tosort_water(self.file_path, self.filter_option)
            
            # 2. Define Step Order based on tab type
            if self.tab_type == "carbonate":
                step_order = [
                    ("Step 1: Data", lambda: step1_data_carbonate(self.file_path, self.sheet_name)),
                    ("Step 2: To Sort", lambda: (refresh_excel(self.file_path), step2_carbonate())), 
                    ("Step 3: Last 6", lambda: step3_last6_carbonate(self.file_path)),
                    ("Step 4: Pre-Group", lambda: step4_pre_group_carbonate(self.file_path)),
                    ("Step 5: Group", lambda: step5_group_carbonate(self.file_path)),
                    ("Step 6: Normalization", lambda: step6_normalization_carbonate(self.file_path)),
                    ("Step 7: Report", lambda: step7_report_carbonate(self.file_path)),
                ]
            else: # self.tab_type == "water"
                step_order = [
                    ("Step 1: Data", lambda: step1_data_water(self.file_path, self.sheet_name)),
                    ("Step 2: To Sort", lambda: (refresh_excel(self.file_path), step2_water())),
                    ("Step 3: Last 6", lambda: step3_last6_water(self.file_path)),
                    ("Step 4: Pre-Group", lambda: (refresh_excel(self.file_path), step4_pre_group_water(self.file_path))),
                    ("Step 5: Group", lambda: (refresh_excel(self.file_path), step5_group_water(self.file_path))),
                    ("Step 6: Normalization", lambda: (refresh_excel(self.file_path), step6_normalization_water(self.file_path))),
                    ("Step 7: Report", lambda: (refresh_excel(self.file_path), step7_report_water(self.file_path))),
                ]
            
            # 3. Execution Logic
            selected_steps = [s for s, checked in self.steps.items() if checked]
            total = len(selected_steps)
            done = 0
            self.log.emit(f"Starting processing for: {os.path.basename(self.file_path)} (Type: {self.tab_type.title()})", "white")
            
            for name, func in step_order:
                # Crucial check: Stop before starting a new step
                if not self._is_running:
                    self.stopped_early.emit()
                    return

                if name in selected_steps:
                    self.progress.emit(done, total, name)
                    try:
                        func()
                        self.log.emit(f"✔ {name} completed", "green")
                        done += 1
                    except Exception as e:
                        self.log.emit(f"✖ {name} failed: {e}", "red")
                        self.progress.emit(done, total, name)
                        break # Stop execution on failure
            
            if self._is_running: # Only emit success if it finished without being stopped
                self.log.emit("✅ All selected steps finished.\n", "green")
                self.progress.emit(total, total, "done")

        except Exception as e:
            self.log.emit(f"Unexpected error: {e}", "red")
            self.progress.emit(done, total, name if 'name' in locals() else "Error")

# ---------------- Password Popup ----------------
class PasswordPopup(QDialog):
    SALT = bytes.fromhex('bae6921f6c798a23813acff5e049d11f')
    HASH = '66de83fdf9831393817c2ce974b05b033b2d723384b9c87e7ac8a2a7a92b732e'

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Unlock Advanced Settings")
        self.setFixedSize(300, 120)
        self.layout = QVBoxLayout(self)
        self.password_correct = False
        
        self.label = QLabel("Enter Password to Unlock:")
        self.layout.addWidget(self.label)
        
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.layout.addWidget(self.password_input)
        
        self.button = QPushButton("Unlock")
        self.button.clicked.connect(self.check_password)
        self.layout.addWidget(self.button)

    def check_password(self):
        # 1. Get the user's plain text input
        input_password = self.password_input.text()
        
        # 2. Hash the user's input using the stored salt and SHA-256
        # Note: input_password must be encoded to bytes before hashing.
        input_hash = hashlib.sha256(self.SALT + input_password.encode('utf-8')).hexdigest()
        
        # 3. Compare the newly generated hash with the stored hash
        if input_hash == self.HASH:
            self.password_correct = True
            self.accept()
        else:
            QMessageBox.critical(self, "Error", "Incorrect Password.")
            self.password_input.clear()
            self.password_input.setFocus()
            
# ---------------- Advanced Settings Tab (No change) ----------------
class AdvancedSettingsTab(QWidget):
    def __init__(self):
        super().__init__()
        
        # 1. Initialize the main layout for the QWidget
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        # 2. Build the UI inside a GroupBox for consistent styling
        self._create_settings_ui()
        
    def _add_divider(self, layout):
        """Helper function for adding a horizontal line."""
        divider = QFrame()
        divider.setFrameShape(QFrame.Shape.HLine)
        divider.setFrameShadow(QFrame.Shadow.Sunken)
        layout.addWidget(divider)

    def _create_settings_ui(self):
        """Generates the Advanced Settings UI inside a QGroupBox."""
        
        # 1. Create the GroupBox for styling (similar to "Carbonate Steps")
        settings_group = QGroupBox("Advanced Settings")
        group_layout = QVBoxLayout()
        settings_group.setLayout(group_layout)
        
        # --- UI Elements ---
        
        # Title (Moved inside the group_layout to appear first)
        title = QLabel("STDEV Threshold Configuration")
        # Note: You might want to adjust this style if it conflicts with QGroupBox title style
        title.setStyleSheet("font-size: 14px; font-weight: bold; margin-bottom: 5px;")
        group_layout.addWidget(title)

        # Horizontal divider
        self._add_divider(group_layout)

        # Settings Group: STDEV_THRESHOLD (Re-using existing QHBoxLayout logic)
        setting_layout = QHBoxLayout()
        setting_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)

        self.label_stdev = QLabel("STDEV_THRESHOLD:")
        setting_layout.addWidget(self.label_stdev)

        self.input_stdev = QLineEdit()
        self.input_stdev.setFixedWidth(100)
        
        # Initialize with the current value from settings.py using the getter
        self.input_stdev.setText(str(settings.get_setting("STDEV_THRESHOLD")))
        setting_layout.addWidget(self.input_stdev)

        self.save_btn = QPushButton("Save")
        self.save_btn.setFixedWidth(80)
        self.save_btn.clicked.connect(self.save_stdev_threshold)
        setting_layout.addWidget(self.save_btn)

        # Add a stretch to the QHBoxLayout to push content left
        setting_layout.addStretch(1) 
        
        # Add the setting row to the group's vertical layout
        group_layout.addLayout(setting_layout)
        
        # Add a stretch to the group's layout to push all content to the top of the group
        group_layout.addStretch(1)

        # --- Final Assembly ---
        
        # Add the GroupBox to the main tab layout
        self.layout.addWidget(settings_group)
        
        # Add a stretch to the tab's main layout to push the GroupBox to the top of the tab
        self.layout.addStretch(1)

    def save_stdev_threshold(self):
        """Attempts to save the new STDEV_THRESHOLD value."""
        new_value_str = self.input_stdev.text()
        
        success, message = settings.set_setting("STDEV_THRESHOLD", new_value_str)

        if success:
            QMessageBox.information(
                self, 
                "Success", 
                f"STDEV_THRESHOLD updated successfully.\n{message}"
            )
            # Ensure the input box shows the newly set (validated) value
            self.input_stdev.setText(str(settings.get_setting("STDEV_THRESHOLD")))
        else:
            QMessageBox.critical(
                self, 
                "Error", 
                f"Failed to update STDEV_THRESHOLD.\n{message}"
            )

# ---------------- Main GUI ----------------
class DataToolApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MRSI - Data Normalization Tool")
        self.setMinimumSize(900, 700)
        self.file_path = None
        self.thread = None
        self.dark_mode = False
        self.is_locked = True
        
        self.active_tab_widget = None
        
        self.carbonate_tab = CarbonateTab()
        self.water_tab = WaterTab()
        self.advanced_settings_tab = AdvancedSettingsTab() # Instantiated from the merged class
        
        self.light_stylesheet = """
QTabBar { background: #ECECEC; padding: 4px; }
QWidget { background-color: #FAFAFA; font-family: Arial; color: #202124; }
QPushButton { background-color: #4CAF50; color: white; border-radius: 6px; padding: 7px 12px; font-size: 13px; }
QPushButton#stopBtn { background-color: #F44336; }
QPushButton#stopBtn:hover { background-color: #D32F2F; }
QPushButton#clearBtn { background-color: #9E9E9E; }
QPushButton#clearBtn:hover { background-color: #757575; }
QPushButton[flat="true"] { background: transparent; color: #333; border: none; }
QPushButton:hover { background-color: #45A049; }
QGroupBox { border: 1px solid #DADCE0; border-radius: 6px; margin-top: 12px; padding: 12px; font-weight: bold; background-color: #F7F7F7; }
QGroupBox#fileGroup { font-size: 14px; }
QLabel#titleLabel { font-size: 18px; font-weight: bold; color: #111; }
QLabel#subtitleLabel { font-size: 14px; color: #333; }
QTextEdit { background-color: #202124; color: #E8EAED; border-radius: 8px; padding: 8px; }
QProgressBar { border: 1px solid #AAA; border-radius: 6px; height: 18px; }
QProgressBar::chunk { background-color: #4CAF50; border-radius: 6px; }
QTabBar::tab { background: #ECECEC; border: 1px solid #D0D0D0; padding: 8px 16px; border-top-left-radius: 6px; border-top-right-radius: 6px; margin-right: 6px; }
QTabBar::tab:selected { background: white; border-bottom: 0px; }
QTabBar::tab:!selected { margin-top: 4px; }
QTabBar::tab:first { margin-left: 6px; }
QLabel#fileInfo { font-size: 13px; color: #111; }
QLabel#noFile { font-size: 13px; color: #777; }
"""
        self.dark_stylesheet = """
QTabBar { background: #2A2B2E; padding: 4px; }
QWidget { background-color: #1E1F22; font-family: Arial; color: #E8EAED; }
QPushButton { background-color: #2E7D32; color: white; border-radius: 6px; padding: 7px 12px; font-size: 13px; }
QPushButton#stopBtn { background-color: #B71C1C; }
QPushButton#stopBtn:hover { background-color: #880E4F; }
QPushButton#clearBtn { background-color: #424242; }
QPushButton#clearBtn:hover { background-color: #616161; }
QPushButton[flat="true"] { background: transparent; color: #DDD; border: none; }
QPushButton:hover { background-color: #376b2e; }
QGroupBox { border: 1px solid #333; border-radius: 6px; margin-top: 12px; padding: 12px; font-weight: bold; background-color: #2A2B2E; }
QGroupBox#fileGroup { font-size: 14px; }
QLabel#titleLabel { font-size: 18px; font-weight: bold; color: #FFF; }
QLabel#subtitleLabel { font-size: 14px; color: #DDD; }
QTextEdit { background-color: #0F1112; color: #E8EAED; border-radius: 8px; padding: 8px; }
QProgressBar { border: 1px solid #444; border-radius: 6px; height: 18px; }
QProgressBar::chunk { background-color: #2E7D32; border-radius: 6px; }
QTabBar::tab { background: #2A2B2E; border: 1px solid #333; padding: 8px 16px; border-top-left-radius: 6px; border-top-right-radius: 6px; margin-right: 6px; color: #DDD; }
QTabBar::tab:selected { background: #222; border-bottom: 0px; }
QLabel#fileInfo { font-size: 13px; color: #EEE; }
QLabel#noFile { font-size: 13px; color: #AAA; }
"""
        self.setStyleSheet(self.light_stylesheet)
        
        self.init_ui()
        self.update_lock_state()

    def init_ui(self):
        root = QVBoxLayout(self)
        self.change_btn = QPushButton("Change File")
        self.remove_btn = QPushButton("Remove")
        self.open_file_btn = QPushButton("📄 Open File")
        self.open_folder_btn = QPushButton("📁 Open Folder")
        
        # ---- Header ----
        header = QHBoxLayout()
        header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title = QLabel("McMaster Research Group for Stable Isotopologues")
        title.setObjectName("titleLabel")
        title.setFont(QFont("Arial", 20, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        header.addWidget(title)
        root.addLayout(header)
        
        subtitle = QLabel("Data Normalization Tool")
        subtitle.setObjectName("subtitleLabel")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setFont(QFont("Arial", 14))
        root.addWidget(subtitle)
        
        # ---- Lock Button ----
        self.lock_btn = QPushButton("🔒", self)
        self.lock_btn.setProperty("flat", True)
        self.lock_btn.setFixedSize(36, 36)
        self.lock_btn.setStyleSheet("""
            QPushButton {
                background-color: #3f51b5;
                color: white;
                border-radius: 18px;
                font-size: 18px;
                padding: 4px 0;
            }
            QPushButton:hover { background-color: #303f9f; }
        """)
        self.lock_btn.clicked.connect(self.toggle_advanced_settings_lock)
        
        # ---- Menu Button (floating, top-right) ----
        self.menu_btn = QPushButton("≡", self)
        self.menu_btn.setProperty("flat", True)
        self.menu_btn.setFixedSize(36, 36)
        self.menu_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 18px;
                font-size: 18px;
                padding: 4px 0;
            }
            QPushButton:hover { background-color: #45A049; }
        """)
        self.menu_btn.clicked.connect(self.show_menu)
        
        QTimer.singleShot(0, self.position_header_buttons)
        
        # ---- File selection box ----
        file_box = QGroupBox("File Selection")
        file_box.setObjectName("fileGroup")
        fl = QVBoxLayout()
        file_info_row = QHBoxLayout()
        self.file_label = QLabel()
        self.file_label.setObjectName("noFile")
        self.set_no_file_label()
        file_info_row.addWidget(self.file_label)
        file_info_row.addStretch()
        self.change_btn.setFixedWidth(110)
        self.change_btn.clicked.connect(self.browse_file)
        self.change_btn.hide()
        self.remove_btn.setFixedWidth(80)
        self.remove_btn.clicked.connect(self.remove_file)
        self.remove_btn.hide()
        file_info_row.addWidget(self.change_btn)
        file_info_row.addWidget(self.remove_btn)
        fl.addLayout(file_info_row)
        divider = QFrame()
        divider.setFrameShape(QFrame.Shape.HLine)
        divider.setFrameShadow(QFrame.Shadow.Sunken)
        fl.addWidget(divider)
        btn_row = QHBoxLayout()
        self.browse_btn = QPushButton("📂 Browse File")
        self.browse_btn.setFixedWidth(140)
        self.browse_btn.clicked.connect(self.browse_file)
        self.open_file_btn.setFixedWidth(120)
        self.open_folder_btn.setFixedWidth(120)
        for b in [self.open_file_btn, self.open_folder_btn]:
            b.hide()
        self.open_file_btn.clicked.connect(self.open_file)
        self.open_folder_btn.clicked.connect(self.open_folder)
        btn_row.addWidget(self.browse_btn)
        btn_row.addWidget(self.open_file_btn)
        btn_row.addWidget(self.open_folder_btn)
        btn_row.addStretch()
        fl.addLayout(btn_row)
        file_box.setLayout(fl)
        root.addWidget(file_box)
        
        # ---- Tabs and Content Frame ----
        tabs_and_content = QVBoxLayout()
        self.tab_bar = QTabBar()
        self.tab_bar.addTab("Carbonate")
        self.tab_bar.addTab("Water")
        self.advanced_tab_index = -1
        
        self.tab_bar.setExpanding(False)
        self.tab_bar.setMovable(False)
        self.tab_bar.setDrawBase(False)
        self.tab_bar.currentChanged.connect(self.on_tab_changed)
        tabs_and_content.addWidget(self.tab_bar)
        
        self.content_frame = QFrame()
        self.content_frame.setObjectName("contentPane")
        self.content_frame.setFrameShape(QFrame.Shape.StyledPanel)
        
        self.content_frame_layout = QVBoxLayout(self.content_frame)
        self.content_frame_layout.setContentsMargins(8, 8, 8, 8)
        self.content_frame_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        tabs_and_content.addWidget(self.content_frame)
        root.addLayout(tabs_and_content)
        
        self.on_tab_changed(0)
        
        # ---- Progress Bar ----
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.hide()
        root.addWidget(self.progress)
        
        # ---- Log Box ----
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.hide()
        root.addWidget(self.log_box)
        
        # ---- Run/Stop/Clear Button Container (UPDATED) ----
        self.button_container = QHBoxLayout()
        self.run_btn = QPushButton("▶ Run Selected Steps")
        self.run_btn.clicked.connect(self.run_steps)
        
        self.stop_btn = QPushButton("◼ Stop")
        self.stop_btn.setObjectName("stopBtn")
        self.stop_btn.clicked.connect(self.stop_steps)
        self.stop_btn.hide() # Initially hidden
        
        self.clear_log_btn = QPushButton("✖ Clear Log")
        self.clear_log_btn.setObjectName("clearBtn")
        self.clear_log_btn.clicked.connect(self.clear_log)
        
        self.button_container.addWidget(self.run_btn)
        self.button_container.addWidget(self.stop_btn)
        self.button_container.addWidget(self.clear_log_btn)
        root.addLayout(self.button_container)
    
    def clear_log(self):
        """Clears the log box and resets the related UI elements."""
        self.log_box.clear()
        self.log_box.hide()
        self.progress.setValue(0)
        self.progress.hide()
        # Reset size if it was expanded for the log
        self.resize(self.width(), 700)
        self.run_btn.show()
        self.stop_btn.hide()

    def position_header_buttons(self):
        self.menu_btn.move(self.width() - 50, 10)
        self.lock_btn.move(self.width() - 50 - 36 - 8, 10)

    def update_lock_state(self):
        """Updates the lock button appearance and tab bar content."""
        if self.is_locked:
            self.lock_btn.setText("🔒")
            self.lock_btn.setStyleSheet("""
                QPushButton { background-color: #808080; color: white; border-radius: 18px; font-size: 18px; padding: 4px 0; }
                QPushButton:hover { background-color: #808080; }
            """)
            if self.advanced_tab_index != -1:
                # Switch away from Advanced Settings if it's currently active
                if self.tab_bar.currentIndex() == self.advanced_tab_index:
                    self.tab_bar.setCurrentIndex(0)
                self.tab_bar.removeTab(self.advanced_tab_index)
                self.advanced_tab_index = -1
        else:
            self.lock_btn.setText("🔓")
            self.lock_btn.setStyleSheet("""
                QPushButton { background-color: #FF9800; color: white; border-radius: 18px; font-size: 18px; padding: 4px 0; }
                QPushButton:hover { background-color: #F57C00; }
            """)
            if self.advanced_tab_index == -1:
                self.advanced_tab_index = self.tab_bar.addTab("Advanced Settings")

    def toggle_advanced_settings_lock(self):
        """Opens the password dialog or locks the settings."""
        if self.is_locked:
            popup = PasswordPopup(self)
            if popup.exec() == QDialog.DialogCode.Accepted:
                if popup.password_correct:
                    self.is_locked = False
                    self.update_lock_state()
                    #QMessageBox.information(self, "Unlocked", "Advanced Settings tab is now visible.")
                    if self.advanced_tab_index != -1:
                        self.tab_bar.setCurrentIndex(self.advanced_tab_index)
        else:
            self.is_locked = True
            self.update_lock_state()
            #QMessageBox.information(self, "Locked", "Advanced Settings tab is now hidden.")
            
    # In gui.py or main.py (where DataToolApp is defined)

    def on_tab_changed(self, index):
        
        # --- PHASE 1: Clean Up the Content Frame ---
        # Iterate over all items currently in the content frame layout
        # and remove them to ensure a clean slate.
        while self.content_frame_layout.count():
            item = self.content_frame_layout.takeAt(0)
            if item.widget():
                item.widget().hide()
            # Note: takeAt(0) also removes stretch items

        # --- PHASE 2: Select and Add the New Tab ---
        if index == 0:
            self.active_tab_widget = self.carbonate_tab
        elif index == 1:
            self.active_tab_widget = self.water_tab
        elif index == self.advanced_tab_index and not self.is_locked:
            self.active_tab_widget = self.advanced_settings_tab
        else:
            self.active_tab_widget = self.carbonate_tab
            self.tab_bar.setCurrentIndex(0)

        # Add the active tab widget
        self.content_frame_layout.addWidget(self.active_tab_widget)
        self.active_tab_widget.show()
        
        # --- PHASE 3: Add a Single Stretch (Critical) ---
        # Add a single stretch item to the END of the layout.
        # This will push the 'active_tab_widget' to the top.
        self.content_frame_layout.addStretch(1)
        
    def show_menu(self):
        menu = QMenu()
        about_action = menu.addAction("About")
        about_action.triggered.connect(self.show_about)
        self.toggle_dm_action = menu.addAction("Toggle Dark Mode")
        self.toggle_dm_action.triggered.connect(self.toggle_dark_mode)
        menu.exec(self.mapToGlobal(self.menu_btn.rect().bottomLeft()))

    def show_about(self):
        python_version = sys.version.split()[0]
        QMessageBox.information(
            self, "About",
            f"""McMaster Research Group for Stable Isotopologues
Data Normalization Tool
Required Python Version: {python_version}
© 2025 McMaster University
Developer: Ibrahim Parvez"""
        )

    def toggle_dark_mode(self):
        self.dark_mode = not self.dark_mode
        if self.dark_mode:
            self.setStyleSheet(self.dark_stylesheet)
            self.carbonate_tab.setStyleSheet(self.dark_stylesheet)
            self.water_tab.setStyleSheet(self.dark_stylesheet)
            self.advanced_settings_tab.setStyleSheet(self.dark_stylesheet)
        else:
            self.setStyleSheet(self.light_stylesheet)
            self.carbonate_tab.setStyleSheet(self.light_stylesheet)
            self.water_tab.setStyleSheet(self.light_stylesheet)
            self.advanced_settings_tab.setStyleSheet(self.light_stylesheet)
    
    # ... (File handling methods remain the same) ...
    def set_no_file_label(self):
        self.file_label.setText("<i>No file selected</i>")
        self.file_label.setObjectName("noFile")
        self.change_btn.hide()
        self.remove_btn.hide()
        self.open_file_btn.hide()
        self.open_folder_btn.hide()

    def set_file_label(self, path):
        fname = os.path.basename(path)
        self.file_label.setText(f"<b>File selected:</b> {fname}")
        self.file_label.setObjectName("fileInfo")
        self.change_btn.show()
        self.remove_btn.show()
        self.open_file_btn.show()
        self.open_folder_btn.show()

    def remove_file(self):
        self.file_path = None
        self.set_no_file_label()

    def browse_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel files (*.xlsx *.xls)"
        )
        if not path:
            return

        if path.lower().endswith(".xls"):
            reply = QMessageBox.question(
                self,
                "Convert to .xlsx?",
                (
                    "The selected file is in the older .xls format.\n\n"
                    "To ensure full compatibility with this tool, it must be "
                    "converted to .xlsx.\n\n"
                    "Would you like to convert it now?"
                ),
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            if reply == QMessageBox.StandardButton.Yes:
                try:
                    new_path = convert_xls_to_xlsx(path)
                    QMessageBox.information(
                        self,
                        "Conversion Complete",
                        f"Successfully converted to:\n{os.path.basename(new_path)}",
                    )
                    self.file_path = new_path
                    self.set_file_label(self.file_path)
                except Exception as e:
                    QMessageBox.critical(
                        self,
                        "Conversion Error",
                        f"Unable to convert file:\n{e}",
                    )
                    self.file_path = None
                    self.set_no_file_label()
                return
            else:
                QMessageBox.warning(
                    self,
                    "Conversion Required",
                    "You must convert the file to .xlsx before running the process.\n\n"
                    "No file was selected.",
                )
                self.file_path = None
                self.set_no_file_label()
                return
        else:
            self.file_path = path
            self.set_file_label(self.file_path)

    def open_file(self):
        if not self.file_path: return
        try:
            if sys.platform.startswith("darwin"): # macOS
                subprocess.run(["open", self.file_path])
            elif os.name == "nt": # Windows
                os.startfile(self.file_path)
            elif os.name == "posix": # Linux
                subprocess.run(["xdg-open", self.file_path])
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Unable to open file:\n{e}")

    def open_folder(self):
        if not self.file_path: return
        folder = os.path.dirname(self.file_path)
        try:
            if sys.platform.startswith("darwin"): # macOS
                subprocess.run(["open", folder])
            elif os.name == "nt": # Windows
                os.startfile(folder)
            elif os.name == "posix": # Linux
                subprocess.run(["xdg-open", folder])
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Unable to open folder:\n{e}")

    # ... (Logging/Progress methods remain the same) ...
    def on_log(self, msg, color):
        cmap = {"red": "#FF6B6B", "green": "#4CAF50", "white": "#E8EAED"}
        text_color = cmap.get(color, "#E8EAED") if self.dark_mode else cmap.get(color, "#202124")
        self.log_box.append(f'<span style="color:{text_color};">{msg}</span>')

    def on_progress(self, done, total, step_name):
        if total == 0:
            return
        val = int((done / total) * 100)
        self.progress.setValue(val)
        if done < total:
            # Change color to indicate ongoing process / not yet finished
            self.progress.setStyleSheet("QProgressBar::chunk { background-color: #3f51b5; }") 
        else:
            self.progress.setStyleSheet("QProgressBar::chunk { background-color: #4CAF50; }")

    def on_thread_done(self):
        """Called when the thread finishes normally (not stopped by user)."""
        self.on_log("Processing complete.", "green")
        self.run_btn.show()
        self.stop_btn.hide()
        self.run_btn.setEnabled(True)
        self.stop_btn.setEnabled(True) # Reset for next run

    def on_thread_stopped(self):
        """Called when the thread finishes due to user-initiated stop."""
        self.on_log("🛑 **Process stopped by user.**\n", "red")
        self.run_btn.show()
        self.stop_btn.hide()
        self.run_btn.setEnabled(True)
        self.stop_btn.setEnabled(True) # Reset for next run
        self.progress.setStyleSheet("QProgressBar::chunk { background-color: #F44336; }") # Red progress bar on stop

    def run_steps(self):
        # 1. Validation
        if not self.file_path:
            QMessageBox.warning(self, "Error", "Please select a file first!")
            return
        if self.file_path.lower().endswith(".xls"):
            QMessageBox.warning(
                self,
                "Incompatible File Format",
                "The selected file is in .xls format.\n\n"
                "Please convert it to .xlsx before running.",
            )
            return
        
        if self.active_tab_widget == self.advanced_settings_tab:
            QMessageBox.warning(self.label_stdev, "Error", "Cannot run from the Advanced Settings tab.")
            return

        # 2. Determine Tab and Collect Parameters from the active tab widget
        current_tab_index = self.tab_bar.currentIndex()
        tab_type = "carbonate" if current_tab_index == 0 else "water"

        try:
            steps, sheet_name, filter_opt = self.active_tab_widget.get_run_parameters()
        except AttributeError:
            QMessageBox.critical(self, "Internal Error", "Could not retrieve parameters from active tab.")
            return

        if not steps or not any(steps.values()):
            QMessageBox.warning(self, "Error", "Please select at least one step to run.")
            return

        # 3. UI State Change for Running
        self.log_box.show()
        self.progress.show()
        self.resize(self.width(), max(self.height(), 950)) 
        self.log_box.clear()
        
        self.run_btn.hide()
        self.stop_btn.show()
        self.stop_btn.setEnabled(True)
        self.progress.setStyleSheet("QProgressBar::chunk { background-color: #3f51b5; }")


        # 4. Start Thread
        self.thread = WorkerThread(self.file_path, steps, sheet_name, filter_opt, tab_type)
        self.thread.log.connect(self.on_log)
        self.thread.progress.connect(self.on_progress)
        self.thread.finished.connect(self.on_thread_done)
        self.thread.stopped_early.connect(self.on_thread_stopped) # Connect new stop signal
        self.thread.start()

    def stop_steps(self):
        """Sends the stop signal to the WorkerThread and updates UI."""
        if self.thread and self.thread.isRunning():
            self.thread.stop()
            self.stop_btn.setEnabled(False) # Disable immediately to prevent spamming
            self.on_log("Stopping process... will complete the current step.", "red")


    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, "menu_btn"):
            self.position_header_buttons()
            
if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = DataToolApp()
    win.show()
    sys.exit(app.exec())