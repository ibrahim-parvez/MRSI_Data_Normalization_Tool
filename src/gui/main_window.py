import sys
import os
import subprocess
import time
import pandas as pd
import xlwings as xw
import hashlib
import base64

from PyQt6.QtGui import QFont, QIcon, QCursor, QPainter, QColor, QPen, QAction, QKeySequence, QPixmap, QImage, QDesktopServices
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QFileDialog, QTabWidget, QTextEdit, QCheckBox,
    QLineEdit, QComboBox, QGroupBox, QMessageBox, QMenu, QProgressBar, QFrame,
    QSizePolicy, QSpacerItem, QGridLayout, QTabBar, QDialog, QScrollArea, QButtonGroup, 
    QRadioButton, QListWidget, QAbstractItemView, QTableWidget, QTableWidgetItem, QHeaderView, QLayout,
    QToolTip, QStyleOptionGroupBox, QProgressDialog, QLabel, QStyle, QDoubleSpinBox
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer, QPoint, QRect, QSize, QPropertyAnimation, QEasingCurve, QByteArray, QUrl

# --- Import Settings ---
import utils.settings as settings
from gui.splash import StartupSplashScreen

# ---- Import Current Version and Updater ---- 
from utils.updater import CURRENT_VERSION, AutoUpdater, apply_update_and_restart, UpdateAvailableDialog

# --- Import Tab UIs ---
from gui.tabs.carbonate_tab import CarbonateTab
from gui.tabs.water_tab import WaterTab
from gui.tabs.combine_tab import CombineTab


# ---- Import your step modules ----
# Ensure 'steps' package exists or comment out for UI testing

from processors.carbonate.step1_data import step1_data_carbonate
from processors.carbonate.step2_tosort import step2_tosort_carbonate
from processors.carbonate.step3_last6 import step3_last6_carbonate
from processors.carbonate.step4_pre_group import step4_pre_group_carbonate
from processors.carbonate.step5_group import step5_group_carbonate
from processors.carbonate.step6_normalization import step6_normalization_carbonate
from processors.carbonate.step7_report import step7_report_carbonate
from processors.water.step1_data import step1_data_water
from processors.water.step2_tosort import step2_tosort_water
from processors.water.step3_last6 import step3_last6_water
from processors.water.step4_pre_group import step4_pre_group_water
from processors.water.step5_group import step5_group_water
from processors.water.step6_normalization import step6_normalization_water
from processors.water.step7_report import step7_report_water


# ---- Imports for Version History ----
import shutil
import tempfile
from datetime import datetime

# ---- Import Current Version ---- 
from utils.updater import CURRENT_VERSION

# ---- Import Logo ----
from utils.logo import logo_base64

# ---------------- Utility: XLS → XLSX ----------------
def convert_xls_to_xlsx(file_path):
    new_path = os.path.splitext(file_path)[0] + ".xlsx"
    try:
        df = pd.read_excel(file_path, engine="xlrd")
        with pd.ExcelWriter(new_path, engine="openpyxl") as writer:
            default_sheet_name = "ExportGB2.wke"
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

class DragDropBox(QGroupBox):
    fileDropped = pyqtSignal(str)

    def __init__(self, title, parent=None):
        super().__init__(title, parent)
        self.setAcceptDrops(True)
        self.drag_active = False

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.toLocalFile().lower().endswith(('.xlsx', '.xls')):
                    event.accept()
                    self.drag_active = True
                    self.update() # Trigger repaint for animation
                    return
        event.ignore()

    def dragLeaveEvent(self, event):
        self.drag_active = False
        self.update()

    def dropEvent(self, event):
        self.drag_active = False
        self.update()
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(('.xlsx', '.xls')):
                self.fileDropped.emit(path)
                return

    def paintEvent(self, event):
        # 1. Draw the standard GroupBox look
        super().paintEvent(event)

        # 2. If a file is hovering, draw the "Cool Upload" overlay
        if self.drag_active:
            painter = QPainter(self)
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            
            # Create semi-transparent overlay
            rect = self.contentsRect()
            overlay_color = QColor("#E3F2FD") 
            overlay_color.setAlpha(200) # Transparency
            
            painter.setBrush(overlay_color)
            
            # Dashed Border
            pen = QPen(QColor("#2196F3"))
            pen.setWidth(3)
            pen.setStyle(Qt.PenStyle.DashLine)
            painter.setPen(pen)
            
            # Draw rounded rect
            painter.drawRoundedRect(rect.adjusted(5, 5, -5, -5), 10, 10)
            
            # Draw Text
            painter.setPen(QColor("#0D47A1"))
            font = QFont("Arial", 16, QFont.Weight.Bold)
            painter.setFont(font)
            painter.drawText(rect, Qt.AlignmentFlag.AlignCenter, "📂 Drop Excel File Here")


# ---------------- Worker Thread ----------------
class WorkerThread(QThread):
    log = pyqtSignal(str, str)
    progress = pyqtSignal(int, int, str)
    stopped_early = pyqtSignal()
    
    def __init__(self, file_path, steps, sheet_name, filter_option, tab_type="carbonate"):
        super().__init__()
        self.file_path = file_path
        self.steps = steps
        self.sheet_name = sheet_name
        self.filter_option = filter_option
        self.tab_type = tab_type
        self._is_running = True

    def stop(self):
        self._is_running = False
        
    def run(self):
        try:
            # Lambda functions to wrap step execution
            step2_carbonate = lambda: step2_tosort_carbonate(self.file_path, self.filter_option)
            step2_water = lambda: step2_tosort_water(self.file_path, self.filter_option)
            
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
            else: # water
                step_order = [
                    ("Step 1: Data", lambda: step1_data_water(self.file_path, self.sheet_name)),
                    ("Step 2: To Sort", lambda: (refresh_excel(self.file_path), step2_water())),
                    ("Step 3: Last 6", lambda: step3_last6_water(self.file_path)),
                    ("Step 4: Pre-Group", lambda: (refresh_excel(self.file_path), step4_pre_group_water(self.file_path))),
                    ("Step 5: Group", lambda: (refresh_excel(self.file_path), step5_group_water(self.file_path))),
                    ("Step 6: Normalization", lambda: (refresh_excel(self.file_path), step6_normalization_water(self.file_path))),
                    ("Step 7: Report", lambda: (refresh_excel(self.file_path), step7_report_water(self.file_path))),
                ]
            
            selected_steps = [s for s, checked in self.steps.items() if checked]
            total = len(selected_steps)
            done = 0
            self.log.emit(f"Starting processing for: {os.path.basename(self.file_path)} (Type: {self.tab_type.title()})", "white")
            
            for name, func in step_order:
                if not self._is_running:
                    self.stopped_early.emit()
                    return

                if name in selected_steps:
                    self.progress.emit(done, total, name)
                    
                    # --- NEW: Log that the step is actively running ---
                    self.log.emit(f"▶  Running {name}...", "white") 
                    
                    try:
                        func()
                        # --- Log success ---
                        self.log.emit(f"✔  {name} Completed", "green")
                        done += 1
                    except Exception as e:
                        self.log.emit(f"✖  {name} Failed: {e}", "red")
                        self.progress.emit(done, total, name)
                        break 
            
            if self._is_running:
                self.log.emit("-" * 50, "white")
                self.log.emit("✅ All selected steps finished.\n", "green")
                self.progress.emit(total, total, "done")

        except Exception as e:
            self.log.emit(f"Unexpected error: {e}", "red")
            self.progress.emit(done, total, "Error")

# ---------------- Password Popup ----------------
class PasswordPopup(QDialog):
    SALT = bytes.fromhex('2a736bd429ce86a48d19d4ae02718995')
    HASH = 'cf2887912b623838bce381c6be8f150fa7961ee203b36e0a05a88dba8b06966b'

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
        input_password = self.password_input.text()
        input_hash = hashlib.sha256(self.SALT + input_password.encode('utf-8')).hexdigest()
        if input_hash == self.HASH:
            self.password_correct = True
            self.accept()
        else:
            QMessageBox.critical(self, "Error", "Incorrect Password.")
            self.password_input.clear()
            self.password_input.setFocus()

class AboutDialog(QDialog):
    """A custom About dialog with a base64 logo and an inline update checker."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About MRSI DNT")
        self.setFixedSize(350, 320)
        
        # Match the main window's theme
        self.dark_mode = getattr(parent, 'dark_mode', False)
        if self.dark_mode:
            self.setStyleSheet("QDialog { background-color: #2A2B2E; color: #E8EAED; } QLabel { color: #E8EAED; }")
            btn_hover = "#444444"
        else:
            self.setStyleSheet("QDialog { background-color: #FAFAFA; color: #202124; } QLabel { color: #202124; }")
            btn_hover = "#E0E0E0"

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # 1. Base64 Logo
        self.logo_label = QLabel()
        self.logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        try:
            image_data = base64.b64decode(logo_base64)
            image = QImage.fromData(image_data)
            pixmap = QPixmap.fromImage(image).scaled(
                100, 100, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation
            )
            self.logo_label.setPixmap(pixmap)
        except Exception as e:
            # Fallback if the base64 string is invalid
            self.logo_label.setText("MRSI")
            self.logo_label.setFont(QFont("Arial", 24, QFont.Weight.Bold))
            self.logo_label.setStyleSheet("color: #7A003C;")
        
        layout.addWidget(self.logo_label)
        layout.addSpacing(10)
        
        # 2. Titles
        title = QLabel("McMaster Research Group\nfor Stable Isotopologues")
        title.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)
        
        subtitle = QLabel("Data Normalization Tool")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(subtitle)
        
        layout.addSpacing(15)
        
        # 3. Version + Refresh Button Row
        version_layout = QHBoxLayout()
        version_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        version_label = QLabel(f"<b>Version:</b> {CURRENT_VERSION}")
        version_label.setFont(QFont("Arial", 10))
        version_layout.addWidget(version_label)
        
        # --- FIXED: Using a native system icon instead of an emoji ---
        self.refresh_btn = QPushButton()
        refresh_icon = self.style().standardIcon(QStyle.StandardPixmap.SP_BrowserReload)
        self.refresh_btn.setIcon(refresh_icon)
        # -------------------------------------------------------------
        
        self.refresh_btn.setFixedSize(26, 26)
        self.refresh_btn.setToolTip("Check for Updates")
        self.refresh_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.refresh_btn.setStyleSheet(f"""
            QPushButton {{ 
                background-color: transparent; 
                border: 1px solid #888; 
                border-radius: 13px; 
            }}
            QPushButton:hover {{ 
                background-color: {btn_hover}; 
                border-color: #555; 
            }}
        """)
        
        # Connect the button to the parent's update method
        self.refresh_btn.clicked.connect(self._on_refresh_clicked)
        
        version_layout.addWidget(self.refresh_btn)
        layout.addLayout(version_layout)
        
        layout.addSpacing(15)
        
        # 4. Footer & Links
        footer = QLabel(
            "Developer: <a href='https://www.linkedin.com/in/ibrahim-parvez' style='color: #2196F3;'>Ibrahim Parvez</a><br><br>"
            "Found a bug? Email: <a href='mailto:parvezi@mcmaster.ca' style='color: #2196F3;'>parvezi@mcmaster.ca</a>"
        )
        footer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        footer.setOpenExternalLinks(True) # Makes the HTML links clickable
        layout.addWidget(footer)
        
        layout.addStretch()

    def _on_refresh_clicked(self):
        """Closes the about box and triggers the update check."""
        self.accept()
        if hasattr(self.parent(), 'check_for_updates'):
            self.parent().check_for_updates()

# ---------------- Custom Layouts ----------------
class FlowLayout(QLayout):
    """Standard FlowLayout to wrap items to the next line."""
    def __init__(self, parent=None, margin=0, spacing=-1):
        super().__init__(parent)
        if parent is not None:
            self.setContentsMargins(margin, margin, margin, margin)
        self.setSpacing(spacing)
        self.itemList = []

    def __del__(self):
        item = self.takeAt(0)
        while item:
            item = self.takeAt(0)

    def addItem(self, item):
        self.itemList.append(item)

    def count(self):
        return len(self.itemList)

    def itemAt(self, index):
        if 0 <= index < len(self.itemList):
            return self.itemList[index]
        return None

    def takeAt(self, index):
        if 0 <= index < len(self.itemList):
            return self.itemList.pop(index)
        return None

    def expandingDirections(self):
        return Qt.Orientation(0)

    def hasHeightForWidth(self):
        return True

    def heightForWidth(self, width):
        height = self._do_layout(QRect(0, 0, width, 0), True)
        return height

    def setGeometry(self, rect):
        super().setGeometry(rect)
        self._do_layout(rect, False)

    def sizeHint(self):
        return self.minimumSize()

    def minimumSize(self):
        size = QSize()
        for item in self.itemList:
            size = size.expandedTo(item.minimumSize())
        size += QSize(2 * self.contentsMargins().top(), 2 * self.contentsMargins().top())
        return size

    def _do_layout(self, rect, test_only):
        x, y = rect.x(), rect.y()
        line_height = 0
        spacing = self.spacing()

        for item in self.itemList:
            wid = item.widget()
            space_x = spacing + wid.style().layoutSpacing(QSizePolicy.ControlType.PushButton, QSizePolicy.ControlType.PushButton, Qt.Orientation.Horizontal)
            space_y = spacing + wid.style().layoutSpacing(QSizePolicy.ControlType.PushButton, QSizePolicy.ControlType.PushButton, Qt.Orientation.Vertical)
            
            next_x = x + item.sizeHint().width() + space_x
            if next_x - space_x > rect.right() and line_height > 0:
                x = rect.x()
                y = y + line_height + space_y
                next_x = x + item.sizeHint().width() + space_x
                line_height = 0

            if not test_only:
                item.setGeometry(QRect(QPoint(x, y), item.sizeHint()))

            x = next_x
            line_height = max(line_height, item.sizeHint().height())

        return y + line_height - rect.y()

# ---------------- Advanced Settings Widgets ----------------

class SlopeGroupWidget(QFrame):
    def __init__(self, index, available_materials, selected_materials, parent_widget):
        super().__init__()
        self.parent_widget = parent_widget 
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setObjectName("slopeGroupFrame") # Tells the global CSS to style this container
        
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(5, 5, 5, 5)
        
        header_row = QHBoxLayout()
        lbl = QLabel(f"<b>Instance {index + 1}:</b>")
        lbl.setObjectName("slopeHeader") # Tells the global CSS to style this text
        header_row.addWidget(lbl)
        header_row.addStretch()
        
        # We keep the delete button's style local because it's a specific, red danger button
        self.del_btn = QPushButton("−") 
        self.del_btn.setFixedSize(24, 24)
        self.del_btn.setToolTip("Remove this group")
        self.del_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.del_btn.setStyleSheet("""
            QPushButton { background-color: #ff4d4d; color: white; border: none; border-radius: 4px; font-weight: bold; font-size: 16px; padding: 0px; }
            QPushButton:hover { background-color: #d32f2f; }
        """)
        self.del_btn.clicked.connect(self._delete_self)
        header_row.addWidget(self.del_btn)
        
        self.layout.addLayout(header_row)
        
        self.chk_container = QWidget()
        self.chk_container.setStyleSheet("border: none;")
        self.chk_layout = FlowLayout(self.chk_container) 
        self.checkboxes = {}
        
        self.update_available_materials(available_materials, selected_materials)
        self.layout.addWidget(self.chk_container)

    def update_available_materials(self, available_materials, selected_materials=None):
        if selected_materials is None:
            selected_materials = self.get_selected_materials()
            
        while self.chk_layout.count():
            item = self.chk_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self.checkboxes.clear()
        
        for mat_name in available_materials:
            chk = QPushButton(mat_name)
            chk.setCheckable(True)
            chk.setProperty("isMaterialToggle", True) # Tells the global CSS how to style these boxes
            chk.setCursor(Qt.CursorShape.PointingHandCursor)
            
            if mat_name in selected_materials:
                chk.setChecked(True)
                
            chk.toggled.connect(self._on_change)
            self.chk_layout.addWidget(chk)
            self.checkboxes[mat_name] = chk

    def get_selected_materials(self):
        return [name for name, chk in self.checkboxes.items() if chk.isChecked()]

    def _on_change(self):
        self.parent_widget.save_slope_config()

    def _delete_self(self):
        self.parent_widget.remove_slope_group_widget(self)

class MaterialTypeWidget(QWidget):
    """
    A reusable widget that contains the Table and Slope configuration 
    for a specific material type (Carbonate or Water).
    """
    def __init__(self, material_type, headers):
        super().__init__()
        self.material_type = material_type
        self.headers = headers
        self.slope_widgets = []
        self._loading = False 
        
        self.layout = QVBoxLayout(self)
        self._create_table_section()
        self.layout.addSpacing(10)
        self._create_slope_section()
        self.load_data()

    def _create_table_section(self):
        grp = QGroupBox(f"{self.material_type} Reference Materials (RM)")
        l = QVBoxLayout()
        grp.setLayout(l)
        
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(self.headers)
        self.table.setAlternatingRowColors(True)
        # ^ Notice: No setStyleSheet here anymore! It uses the global theme.
        
        header = self.table.horizontalHeader()
        for i in range(7):
            header.setSectionResizeMode(i, QHeaderView.ResizeMode.Stretch)
            
        header.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)
        header.setMinimumSectionSize(30)
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
              
        self.table.setMinimumHeight(200)
        self.table.itemChanged.connect(self._on_table_item_changed)
        l.addWidget(self.table)
        
        btn_layout = QHBoxLayout()
        self.add_btn = QPushButton("Add Row")
        self.add_btn.clicked.connect(self.add_row)
        
        self.remove_btn = QPushButton("Remove Row")
        self.remove_btn.clicked.connect(self.remove_row)
        
        btn_layout.addWidget(self.add_btn)
        btn_layout.addWidget(self.remove_btn)
        btn_layout.addStretch()
        l.addLayout(btn_layout)
        
        self.layout.addWidget(grp)

    def _create_slope_section(self):
        grp = QGroupBox(f"{self.material_type} Slope and Intercept Groups")
        l = QVBoxLayout()
        grp.setLayout(l)
        
        self.slope_container = QWidget()
        self.slope_layout = QVBoxLayout(self.slope_container)
        self.slope_layout.setContentsMargins(0,0,0,0)
        l.addWidget(self.slope_container)
        
        self.add_slope_btn = QPushButton("Add Slope Instance")
        self.add_slope_btn.clicked.connect(self.add_slope_group)
        l.addWidget(self.add_slope_btn)
        self.layout.addWidget(grp)

    def load_data(self):
        self._loading = True
        mats = settings.get_setting("REFERENCE_MATERIALS", sub_key=self.material_type)
        self.table.setRowCount(0)
        if mats:
            for mat in mats: self._insert_table_row(mat)
        self.refresh_slope_ui()
        self._loading = False

    def _insert_table_row(self, mat_data=None):
        if mat_data is None: mat_data = {}
        row = self.table.rowCount()
        self.table.insertRow(row)
        keys = ["col_c", "col_d", "col_e", "col_f", "col_g", "col_h"]
        
        # UNIVERSAL HIGHLIGHT: A transparent gray that works beautifully in both Light and Dark mode!
        highlight_color = QColor(128, 128, 128, 40) 
        
        for i, key in enumerate(keys):
            item = QTableWidgetItem(str(mat_data.get(key, "")))
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            
            if i == 0:
                item.setBackground(highlight_color)
                item.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
            
            self.table.setItem(row, i, item)
            
        combo = QComboBox()
        colors = ["black", "green", "red", "darkblue", "lightblue", "orange", "purple", "brown", "magenta", "teal", "blue", "cyan"]
        combo.addItems(colors)
        combo.setCurrentText(mat_data.get("color", "black"))
        combo.currentTextChanged.connect(lambda: self._save_table_data())
        self.table.setCellWidget(row, 6, combo)

    def refresh_slope_ui(self):
        for i in reversed(range(self.slope_layout.count())): 
            w = self.slope_layout.itemAt(i).widget()
            if w: w.setParent(None)
        self.slope_widgets.clear()
        
        slope_groups = settings.get_setting("SLOPE_INTERCEPT_GROUPS", sub_key=self.material_type)
        available = settings.get_reference_names(self.material_type)
        if slope_groups:
            for i, group in enumerate(slope_groups):
                w = SlopeGroupWidget(i, available, group, parent_widget=self)
                self.slope_layout.addWidget(w)
                self.slope_widgets.append(w)

    def add_row(self):
        self._loading = True
        self._insert_table_row()
        self._loading = False
        self._save_table_data()

    def remove_row(self):
        curr = self.table.currentRow()
        if curr >= 0:
            self.table.removeRow(curr)
            self._save_table_data()

    def add_slope_group(self):
        available = settings.get_reference_names(self.material_type)
        idx = len(self.slope_widgets)
        w = SlopeGroupWidget(idx, available, [], parent_widget=self)
        self.slope_layout.addWidget(w)
        self.slope_widgets.append(w)
        self.save_slope_config()

    def remove_slope_group_widget(self, widget):
        self.slope_layout.removeWidget(widget)
        widget.deleteLater()
        if widget in self.slope_widgets:
            self.slope_widgets.remove(widget)
        self.save_slope_config()

    def _on_table_item_changed(self, item):
        if self._loading: return
        self._save_table_data()

    def _save_table_data(self):
        if self._loading: return
        new_data = []
        for r in range(self.table.rowCount()):
            row_data = {
                "col_c": self.table.item(r, 0).text().strip(),
                "col_d": self.table.item(r, 1).text().strip(),
                "col_e": self.table.item(r, 2).text().strip(),
                "col_f": self.table.item(r, 3).text().strip(),
                "col_g": self.table.item(r, 4).text().strip(),
                "col_h": self.table.item(r, 5).text().strip(),
                "color": self.table.cellWidget(r, 6).currentText(),
                "bold": False
            }
            new_data.append(row_data)
        settings.set_setting("REFERENCE_MATERIALS", new_data, sub_key=self.material_type)
        self._refresh_slope_dropdowns()

    def _refresh_slope_dropdowns(self):
        available = settings.get_reference_names(self.material_type)
        for w in self.slope_widgets:
            w.update_available_materials(available)

    def save_slope_config(self):
        if self._loading: return
        new_config = []
        for w in self.slope_widgets:
            sel = w.get_selected_materials()
            if sel: new_config.append(sel)
        settings.set_setting("SLOPE_INTERCEPT_GROUPS", new_config, sub_key=self.material_type)


class InstantTooltipLabel(QLabel):
    def enterEvent(self, event):
        # Explicitly show the tooltip immediately at the global cursor position
        QToolTip.showText(QCursor.pos(), self.toolTip(), self)
        super().enterEvent(event)


class AdvancedSettingsTab(QWidget):
    def __init__(self):
        super().__init__()
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        
        self.main_layout = QVBoxLayout(self)
        
        # Scroll Area
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        # Disable Horizontal Scroll to enforce width constraints
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff) 
        
        self.scroll_content = QWidget()
        self.layout = QVBoxLayout(self.scroll_content)
        self.layout.setContentsMargins(10, 10, 10, 10)
        self.layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        
        self._create_ui()
        
        self.scroll_area.setWidget(self.scroll_content)
        self.main_layout.addWidget(self.scroll_area)
        
    def _create_ui(self):
        self._create_general_config()
        self._create_outlier_settings()
        self._create_calc_logic_section()
        self._create_material_tabs()
        
    def _create_general_config(self):
        group = QGroupBox("General Configuration")
        layout = QVBoxLayout() 
        group.setLayout(layout)
        
        # --- Helper for Info Icon ---
        def create_info_label(tooltip_text):
            lbl = InstantTooltipLabel("ⓘ") 
            lbl.setCursor(Qt.CursorShape.WhatsThisCursor)
            lbl.setToolTip(tooltip_text)
            lbl.setStyleSheet("""
                QLabel {
                    color: #555;
                    font-size: 14px;
                    font-weight: bold;
                    margin-left: 2px;
                    margin-right: 5px;
                }
                QLabel:hover {
                    color: #0078d7; 
                }
            """)
            return lbl

        # --- Row 1: Stdev Threshold ---
        row1 = QHBoxLayout()
        
        # Stack label and subscript vertically
        lbl_layout = QVBoxLayout()
        lbl_layout.setSpacing(0)
        lbl_layout.addWidget(QLabel("<b>Stdev Threshold</b>"))
        lbl_layout.addWidget(QLabel("<small style='color: gray;'>(All Steps)</small>"))
        
        row1.addLayout(lbl_layout)
        row1.addWidget(create_info_label(
            "<b>Standard Deviation Limit</b><br>"
            "Defines the cutoff value for the standard deviation.<br>"
            "Any value above this limit will be highlighted <span style='color:red;'>red</span> in the stdev columns."
        ))
        row1.addWidget(QLabel(":"))
        
        # --- FIXED: Use QDoubleSpinBox but HIDE the broken native arrows ---
        self.input_stdev = QDoubleSpinBox()
        self.input_stdev.setButtonSymbols(QDoubleSpinBox.ButtonSymbols.NoButtons) # Removes the blank arrows
        self.input_stdev.setFixedWidth(50) 
        self.input_stdev.setDecimals(2)    
        self.input_stdev.setSingleStep(0.01) 
        self.input_stdev.setRange(0.0, 100.0) 
        
        # Safely load the setting as a float
        current_stdev = float(settings.get_setting("STDEV_THRESHOLD") or 0.8)
        self.input_stdev.setValue(current_stdev)
        self.input_stdev.editingFinished.connect(self._on_stdev_changed)
        
        row1.addWidget(self.input_stdev)
        
        # --- Create Custom Up/Down Buttons ---
        btn_up = QPushButton("▲")
        btn_down = QPushButton("▼")
        
        for btn in [btn_up, btn_down]:
            btn.setFixedSize(20, 13)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            # Neutral styling that works beautifully in both Light and Dark mode
            btn.setStyleSheet("""
                QPushButton { 
                    background-color: transparent; 
                    color: #888; 
                    border: 1px solid #888; 
                    border-radius: 2px; 
                    font-size: 8px; 
                    padding: 0px;
                }
                QPushButton:hover { 
                    background-color: #888; 
                    color: white; 
                }
            """)
            
        # Hook them up to the SpinBox's built-in math functions
        btn_up.clicked.connect(self.input_stdev.stepUp)
        btn_down.clicked.connect(self.input_stdev.stepDown)
        
        # Stack the buttons vertically next to the input
        spin_btn_layout = QVBoxLayout()
        spin_btn_layout.setSpacing(2)
        spin_btn_layout.setContentsMargins(0, 0, 0, 0)
        spin_btn_layout.addWidget(btn_up)
        spin_btn_layout.addWidget(btn_down)
        
        row1.addLayout(spin_btn_layout)
        row1.addStretch()
        
        # Add the layout directly to the group box layout
        layout.addLayout(row1)

        self.layout.addWidget(group)


    def _create_outlier_settings(self):
        group = QGroupBox("Outlier Settings")
        layout = QVBoxLayout() 
        group.setLayout(layout)
        
        # --- Helper for Info Icon ---
        def create_info_label(tooltip_text):
            lbl = InstantTooltipLabel("ⓘ") 
            lbl.setCursor(Qt.CursorShape.WhatsThisCursor)
            lbl.setToolTip(tooltip_text)
            lbl.setStyleSheet("""
                QLabel {
                    color: #555;
                    font-size: 14px;
                    font-weight: bold;
                    margin-left: 2px;
                    margin-right: 5px;
                }
                QLabel:hover {
                    color: #0078d7; 
                }
            """)
            return lbl

        # --- Row 1: Outlier Threshold (Sigma) ---
        row1 = QHBoxLayout()
        
        lbl_layout1 = QVBoxLayout()
        lbl_layout1.setSpacing(0)
        lbl_layout1.addWidget(QLabel("<b>Outlier Calculation</b>"))
        lbl_layout1.addWidget(QLabel("<small style='color: gray;'>(Steps: Data, Group, Normalization)</small>"))
        
        row1.addLayout(lbl_layout1)
        row1.addWidget(create_info_label(
            "<b>Sigma Threshold (Standard Deviations)</b><br>"
            "Determines how strict the outlier detection is.<br>"
            "<ul>"
            "<li><b>1σ:</b> avg +- std</li>"
            "<li><b>2σ:</b> avg +- 2*std</li>"
            "<li><b>3σ:</b> avg +- 3*std</li>"
            "</ul>"
        ))
        row1.addWidget(QLabel(":"))

        self.bg_sigma = QButtonGroup(self)
        self.rb_1sigma = QRadioButton("1σ")
        self.rb_2sigma = QRadioButton("2σ")
        self.rb_3sigma = QRadioButton("3σ")
        
        self.bg_sigma.addButton(self.rb_1sigma, 1)
        self.bg_sigma.addButton(self.rb_2sigma, 2)
        self.bg_sigma.addButton(self.rb_3sigma, 3)
        
        row1.addWidget(self.rb_1sigma)
        row1.addWidget(self.rb_2sigma)
        row1.addWidget(self.rb_3sigma)
        row1.addStretch()
        layout.addLayout(row1)

        # --- Row 2: Exclusion Logic ---
        row2 = QHBoxLayout()
        
        lbl_layout2 = QVBoxLayout()
        lbl_layout2.setSpacing(0)
        lbl_layout2.addWidget(QLabel("<b>Exclusion Logic</b>"))
        lbl_layout2.addWidget(QLabel("<small style='color: gray;'>(Steps: Data, Group, Normalization)</small>"))
        
        row2.addLayout(lbl_layout2)
        row2.addWidget(create_info_label(
            "<b>How to Handle Outliers</b><br>"
            "<ul>"
            "<li><b>Individual:</b> If Carbon (δ13C) is an outlier but Oxygen (δ18O) is good, keep the Oxygen value.</li>"
            "<li><b>Exclude Row:</b> If <i>either</i> value is an outlier, discard the entire measurement row.</li>"
            "</ul>"
        ))
        row2.addWidget(QLabel(":"))

        self.bg_excl = QButtonGroup(self)
        self.rb_excl_row = QRadioButton("Exclude Entire Row")
        self.rb_excl_ind = QRadioButton("Individual (Keep Valid C or O)")
        
        self.bg_excl.addButton(self.rb_excl_ind)
        self.bg_excl.addButton(self.rb_excl_row)
        
        row2.addWidget(self.rb_excl_ind)
        row2.addWidget(self.rb_excl_row)
        row2.addStretch()
        layout.addLayout(row2)

        # --- Load Initial Settings ---
        curr_sigma = settings.get_setting("OUTLIER_SIGMA") or 2
        if curr_sigma == 1: self.rb_1sigma.setChecked(True)
        elif curr_sigma == 3: self.rb_3sigma.setChecked(True)
        else: self.rb_2sigma.setChecked(True)
        
        curr_excl = settings.get_setting("OUTLIER_EXCLUSION_MODE") or "Individual"
        if curr_excl == "Exclude Row": self.rb_excl_row.setChecked(True)
        else: self.rb_excl_ind.setChecked(True)

        # --- Connect Signals ---
        self.bg_sigma.idClicked.connect(self._on_sigma_changed)
        self.bg_excl.buttonToggled.connect(self._on_excl_mode_changed)

        self.layout.addWidget(group)


    def _create_calc_logic_section(self):
        group = QGroupBox("Data Selection")
        layout = QVBoxLayout()
        group.setLayout(layout)
        
        # --- Helper for Info Icon ---
        def create_info_label(tooltip_text):
            lbl = InstantTooltipLabel("ⓘ") 
            lbl.setCursor(Qt.CursorShape.WhatsThisCursor)
            lbl.setToolTip(tooltip_text)
            lbl.setStyleSheet("""
                QLabel {
                    color: #555;
                    font-size: 14px;
                    font-weight: bold;
                    margin-left: 2px;
                    margin-right: 5px;
                }
                QLabel:hover {
                    color: #0078d7; 
                }
            """)
            return lbl

        # Step 3
        row1 = QHBoxLayout()
        
        lbl_layout1 = QVBoxLayout()
        lbl_layout1.setSpacing(0)
        lbl_layout1.addWidget(QLabel("<b>Measured 𝛅 values</b>"))
        lbl_layout1.addWidget(QLabel("<small style='color: gray;'>(Step 3: Last 6)</small>"))
        
        row1.addLayout(lbl_layout1)
        row1.addWidget(create_info_label(
            "<b>Calculation Mode for Step 3</b><br>"
            "Decides which data is used to calculate the 'Last 6' Averages.<br>"
            "<ul>"
            "<li><b>Last 6:</b> Takes the raw average of the last 6 measurements.</li>"
            "<li><b>Last 6 Outliers Excluded:</b> Removes statistical outliers <i>before</i> calculating the average.</li>"
            "</ul>"
        ))
        row1.addWidget(QLabel(":"))

        self.bg_step3 = QButtonGroup(self)
        self.rb_s3_last6 = QRadioButton("Last 6")
        self.rb_s3_last6_excl = QRadioButton("Last 6 Outliers Excluded")
        self.bg_step3.addButton(self.rb_s3_last6)
        self.bg_step3.addButton(self.rb_s3_last6_excl)
        row1.addWidget(self.rb_s3_last6)
        row1.addWidget(self.rb_s3_last6_excl)
        row1.addStretch()
        layout.addLayout(row1)
        
        # Step 7
        row2 = QHBoxLayout()
        
        lbl_layout2 = QVBoxLayout()
        lbl_layout2.setSpacing(0)
        lbl_layout2.addWidget(QLabel("<b>Average for RM</b>"))
        lbl_layout2.addWidget(QLabel("<small style='color: gray;'>(Step 7: Normalization)</small>"))
        
        row2.addLayout(lbl_layout2)
        row2.addWidget(create_info_label(
            "<b>Normalization Calculation</b><br>"
            "Determines which data points are used to calculate the Average and Standard Deviation for the Reference Materials (RMs) during normalization.<br>"
            "<ul>"
            "<li><b>All Values:</b> Computes the metrics using every measurement, including those flagged as outliers.</li>"
            "<li><b>Outliers Excluded:</b> Computes the metrics using only valid data points, ignoring any measurements flagged as outliers.</li>"
            "</ul>"
        ))
        row2.addWidget(QLabel(":"))

        self.bg_step7 = QButtonGroup(self)
        self.rb_s7_all = QRadioButton("All Values")
        self.rb_s7_outlier = QRadioButton("Outliers Excluded")
        self.bg_step7.addButton(self.rb_s7_all)
        self.bg_step7.addButton(self.rb_s7_outlier)
        row2.addWidget(self.rb_s7_all)
        row2.addWidget(self.rb_s7_outlier)
        row2.addStretch()
        layout.addLayout(row2)
        
        # Load State
        if settings.get_setting("CALC_MODE_STEP3") == "Last 6 Outliers Excl.": self.rb_s3_last6_excl.setChecked(True)
        else: self.rb_s3_last6.setChecked(True)
            
        if settings.get_setting("CALC_MODE_STEP7") == "Outliers Excluded": self.rb_s7_outlier.setChecked(True)
        else: self.rb_s7_all.setChecked(True)

        self.bg_step3.buttonToggled.connect(self._on_calc_mode_changed)
        self.bg_step7.buttonToggled.connect(self._on_calc_mode_changed)

        self.layout.addWidget(group)

    def _create_material_tabs(self):
        self.tabs = QTabWidget()
        
        self.water_widget = MaterialTypeWidget("Water",
                                               ["Water Standards", "Col D", "Col E", "Col F (δ²H)", "Col G (δ¹⁸O SMOW)", "Col H", "Color"])
        self.tabs.addTab(self.water_widget, "Water")

        self.carb_widget = MaterialTypeWidget("Carbonate", 
                                              ["Col C (Name)", "Col D", "Col E", "Col F (d13C)", "Col G (d18O)", "Col H", "Color"])
        self.tabs.addTab(self.carb_widget, "Carbonate")
        
        self.layout.addWidget(self.tabs)

    def _on_stdev_changed(self):
        val = self.input_stdev.value() # Retrieves a float instead of text
        success, msg = settings.set_setting("STDEV_THRESHOLD", val)
        if not success:
            QMessageBox.warning(self, "Invalid Input", msg)
            # Revert to the last known good setting if it fails
            safe_val = float(settings.get_setting("STDEV_THRESHOLD") or 0.8)
            self.input_stdev.setValue(safe_val)

    def _on_calc_mode_changed(self, btn, checked):
        if not checked: return
        val3 = "Last 6 Outliers Excl." if self.rb_s3_last6_excl.isChecked() else "Last 6"
        settings.set_setting("CALC_MODE_STEP3", val3)
        val7 = "Outliers Excluded" if self.rb_s7_outlier.isChecked() else "All Values"
        settings.set_setting("CALC_MODE_STEP7", val7)
    
    def _on_sigma_changed(self, btn_id):
        settings.set_setting("OUTLIER_SIGMA", btn_id)
        
    def _on_excl_mode_changed(self, btn, checked):
        if not checked: return
        mode = "Individual" if self.rb_excl_ind.isChecked() else "Exclude Row"
        settings.set_setting("OUTLIER_EXCLUSION_MODE", mode)

class LongPressButton(QPushButton):
    """
    A QPushButton that detects a long press (2 seconds) separately from a normal click.
    """
    longPressed = pyqtSignal()

    def __init__(self, text="", parent=None):
        super().__init__(text, parent)
        self._timer = QTimer(self)
        self._timer.setInterval(2000)  # Reduced to 2 seconds for better usability
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self._on_long_press_timeout)
        self._long_press_triggered = False

    def mousePressEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self._long_press_triggered = False
            self._timer.start()
        super().mousePressEvent(e)
        e.accept()

    def mouseReleaseEvent(self, e):
        # Stop the timer immediately upon release
        self._timer.stop()
        
        # Only perform a normal click if we haven't triggered the long press yet
        if e.button() == Qt.MouseButton.LeftButton and not self._long_press_triggered:
            super().mouseReleaseEvent(e)
        
        e.accept()

    def _on_long_press_timeout(self):
        self._long_press_triggered = True
        self.longPressed.emit()
        # Visually un-press the button so it doesn't look stuck
        self.setDown(False)

# ---------------- Version History UI ----------------
class VersionHistoryDialog(QDialog):
    def __init__(self, history, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Version History")
        self.setMinimumSize(450, 300)
        self.history = history # List of dicts
        self.selected_path = None
        self.action = None 

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Select a previous version to view or restore:"))
        
        self.list_widget = QListWidget()
        # Show newest backups at the top
        for item in reversed(self.history): 
            self.list_widget.addItem(f"[{item['time']}] {item['desc']}")
        layout.addWidget(self.list_widget)

        btn_layout = QHBoxLayout()
        self.open_btn = QPushButton("📄 Open/View")
        self.restore_btn = QPushButton("↺ Restore This Version")
        self.cancel_btn = QPushButton("Cancel")

        self.open_btn.setStyleSheet("background-color: #2196F3; color: white;")
        self.restore_btn.setStyleSheet("background-color: #FF9800; color: white;")

        btn_layout.addWidget(self.open_btn)
        btn_layout.addWidget(self.restore_btn)
        btn_layout.addWidget(self.cancel_btn)
        layout.addLayout(btn_layout)

        self.open_btn.clicked.connect(self.on_open)
        self.restore_btn.clicked.connect(self.on_restore)
        self.cancel_btn.clicked.connect(self.reject)

    def get_selected_item(self):
        row = self.list_widget.currentRow()
        if row < 0: return None
        return self.history[len(self.history) - 1 - row] # Match reversed list

    def on_open(self):
        item = self.get_selected_item()
        if item:
            self.selected_path = item['path']
            self.action = 'open'
            self.accept()

    def on_restore(self):
        item = self.get_selected_item()
        if item:
            self.selected_path = item['path']
            self.action = 'restore'
            self.accept()

# ---------------- Main GUI ----------------
class DataToolApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MRSI - Data Normalization Tool")
        self.setMinimumSize(730, 825)

        icon_pixmap = QPixmap()
        icon_pixmap.loadFromData(QByteArray.fromBase64(logo_base64.encode('utf-8')))
        self.setWindowIcon(QIcon(icon_pixmap))

        
        self.file_path = None
        self.thread = None
        self.dark_mode = False
        self.is_locked = True
        
        self.active_tab_widget = None
        

        self.progress_anim = None
        
        self.carbonate_tab = CarbonateTab()
        self.water_tab = WaterTab()
        self.combine_tab = CombineTab()
        self.advanced_settings_tab = AdvancedSettingsTab()

        self.light_stylesheet = """
        QWidget { background-color: #FAFAFA; font-family: Arial; color: #202124; }
        QLabel, QCheckBox, QRadioButton { background-color: transparent; }
        
        QTabBar { background: #ECECEC; padding: 4px; }
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
        QListWidget { background-color: #FFFFFF; border: 1px solid #CCC; border-radius: 4px; }
        QMenu { background-color: #FFFFFF; border: 1px solid #CCC; padding: 5px; border-radius: 6px; }
        QMenu::item { padding: 8px 25px; background-color: transparent; color: #333; border-radius: 4px; }
        QMenu::item:selected { background-color: #E0E0E0; }
        
        /* --- NEW: TABLE STYLING --- */
        QTableWidget { gridline-color: #cccccc; background-color: white; alternate-background-color: #f5f5f5; color: #333; }
        QHeaderView::section { background-color: #e0e0e0; color: #333; padding: 4px; border: 1px solid #bfbfbf; font-weight: bold; }
        QTableCornerButton::section { background-color: #e0e0e0; border: 1px solid #bfbfbf; }
        
        /* --- NEW: SLOPE GROUP STYLING --- */
        QFrame#slopeGroupFrame { background-color: #f9f9f9; border: 1px solid #ddd; border-radius: 4px; }
        QLabel#slopeHeader { color: #333; border: none; }
        QPushButton[isMaterialToggle="true"] { background-color: #ffffff; color: #333; border: 1px solid #888; padding: 5px 10px; border-radius: 4px; font-weight: bold; }
        QPushButton[isMaterialToggle="true"]:hover { background-color: #f0f0f0; border: 1px solid #333; }
        QPushButton[isMaterialToggle="true"]:checked { background-color: #005a9e; color: white; border: 1px solid #004080; }

        /* --- NEW: RADIO BUTTON STYLING --- */
        QRadioButton::indicator {
            width: 14px;
            height: 14px;
            border-radius: 8px; /* (14px width + 2px border) / 2 = 8px perfect circle */
            border: 1px solid #AAA;
            background-color: #FFFFFF;
        }
        QRadioButton::indicator:checked {
            border: 1px solid #005a9e;
            background-color: qradialgradient(cx:0.5, cy:0.5, radius:0.5, fx:0.5, fy:0.5,
                stop:0 #005a9e, stop:0.45 #005a9e, stop:0.55 #FFFFFF, stop:1 #FFFFFF);
        }
        QRadioButton::indicator:hover {
            border: 1px solid #005a9e;
        }
        """
        
        self.dark_stylesheet = """
        QWidget { background-color: #1E1F22; font-family: Arial; color: #E8EAED; }
        QLabel, QCheckBox, QRadioButton { background-color: transparent; }
        
        QTabBar { background: #2A2B2E; padding: 4px; }
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
        QListWidget { background-color: #2A2B2E; border: 1px solid #444; border-radius: 4px; color: #EEE; }
        QMenu { background-color: #2A2B2E; border: 1px solid #444; padding: 5px; border-radius: 6px; }
        QMenu::item { padding: 8px 25px; background-color: transparent; color: #EEE; border-radius: 4px; }
        QMenu::item:selected { background-color: #444; }
        
        /* --- NEW: TABLE STYLING --- */
        QTableWidget { gridline-color: #444444; background-color: #1E1F22; alternate-background-color: #2A2B2E; color: #E8EAED; }
        QHeaderView::section { background-color: #333333; color: #E8EAED; padding: 4px; border: 1px solid #444444; font-weight: bold; }
        QTableCornerButton::section { background-color: #333333; border: 1px solid #444444; }
        
        /* --- NEW: SLOPE GROUP STYLING --- */
        QFrame#slopeGroupFrame { background-color: #2A2B2E; border: 1px solid #555; border-radius: 4px; }
        QLabel#slopeHeader { color: #E8EAED; border: none; }
        QPushButton[isMaterialToggle="true"] { background-color: #1E1F22; color: #E8EAED; border: 1px solid #666; padding: 5px 10px; border-radius: 4px; font-weight: bold; }
        QPushButton[isMaterialToggle="true"]:hover { background-color: #333; border: 1px solid #999; }
        QPushButton[isMaterialToggle="true"]:checked { background-color: #2E7D32; color: white; border: 1px solid #1B5E20; }

        /* --- NEW: RADIO BUTTON STYLING --- */
        QRadioButton::indicator {
            width: 14px;
            height: 14px;
            border-radius: 8px;
            border: 1px solid #666;
            background-color: #2A2B2E;
        }
        QRadioButton::indicator:checked {
            border: 1px solid #2E7D32;
            background-color: qradialgradient(cx:0.5, cy:0.5, radius:0.5, fx:0.5, fy:0.5,
                stop:0 #2E7D32, stop:0.45 #2E7D32, stop:0.55 #2A2B2E, stop:1 #2A2B2E);
        }
        QRadioButton::indicator:hover {
            border: 1px solid #2E7D32;
        }
        """
        # --- OS Default Theme Detection ---
        palette = QApplication.palette()
        window_color = palette.color(palette.ColorRole.Window)
        
        # If the OS window background is dark (lightness < 128), default to dark mode
        self.dark_mode = window_color.lightness() < 128
        
        # Apply the detected stylesheet
        if self.dark_mode:
            self.setStyleSheet(self.dark_stylesheet)
        else:
            self.setStyleSheet(self.light_stylesheet)
        # ----------------------------------
        
        self.init_ui()
        
        # --- KEYBOARD SHORTCUT (Robust Fix using QAction) ---
        # We use QAction because it attaches to the window more reliably than QShortcut.
        self.unlock_action = QAction(self)
        
        # Qt automatically maps "Ctrl" to "Command" on macOS.
        # So this single line works for both Windows (Ctrl+Shift+S) and Mac (Cmd+Shift+S).
        self.unlock_action.setShortcut(QKeySequence("Ctrl+Shift+S"))
        
        # Critical: Add the action to the main window so it listens globally
        self.addAction(self.unlock_action)
        
        self.unlock_action.triggered.connect(self.trigger_secret_unlock)

    # ADD TO THE BOTTOM OF def __init__(self):
        self.version_history = []
        self.history_temp_dir = tempfile.mkdtemp(prefix="mrsi_history_")

    # ADD THESE NEW METHODS INSIDE DataToolApp:
    def save_version(self, description, target_path=None):
        """Creates a copy of the current file in the temp directory."""
        path_to_backup = target_path or self.file_path
        if not path_to_backup or not os.path.exists(path_to_backup): return

        timestamp = datetime.now().strftime("%I:%M:%S %p")
        safe_time = datetime.now().strftime("%H%M%S")
        ext = os.path.splitext(path_to_backup)[1]
        base = os.path.basename(path_to_backup).replace(ext, "")
        
        backup_name = f"{base}_{safe_time}_{len(self.version_history)}{ext}"
        backup_path = os.path.join(self.history_temp_dir, backup_name)
        
        try:
            shutil.copy2(path_to_backup, backup_path)
            self.version_history.append({
                'time': timestamp,
                'desc': description,
                'path': backup_path
            })
            self.history_btn.show() # Unhide button once history exists
        except Exception as e:
            self.on_log(f"Warning: Failed to create backup: {e}", "red")

    def reset_history(self):
        """Clears history and deletes temp files when a new file is loaded."""
        self.version_history.clear()
        self.history_btn.hide()
        for filename in os.listdir(self.history_temp_dir):
            filepath = os.path.join(self.history_temp_dir, filename)
            try:
                os.remove(filepath)
            except Exception:
                pass

    def show_history_dialog(self):
        if not self.version_history: return

        dialog = VersionHistoryDialog(self.version_history, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            if dialog.action == 'open':
                try:
                    if sys.platform.startswith("darwin"): subprocess.run(["open", dialog.selected_path])
                    elif os.name == "nt": os.startfile(dialog.selected_path)
                    elif os.name == "posix": subprocess.run(["xdg-open", dialog.selected_path])
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Unable to open file:\n{e}")
            elif dialog.action == 'restore':
                reply = QMessageBox.question(self, "Confirm Restore", 
                                             "Overwrite your current working file with this older version?\n\n(A backup of your current state will be saved just in case)",
                                             QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if reply == QMessageBox.StandardButton.Yes:
                    try:
                        self.save_version("State right before user Restoration") 
                        shutil.copy2(dialog.selected_path, self.file_path)
                        QMessageBox.information(self, "Restored", "File successfully reverted!")
                        self.on_log("File reverted to an older version.", "white")
                    except Exception as e:
                        QMessageBox.critical(self, "Error", f"Failed to restore file:\n{e}")

    def closeEvent(self, event):
        """Ensures the temporary directory is deleted when the app closes."""
        try:
            shutil.rmtree(self.history_temp_dir)
        except Exception:
            pass
        super().closeEvent(event)
    

    def process_selected_file(self, path):
        """Helper to handle file loading for both Browse and Drop events"""
        if not path: return

        # Reset history if they select a brand new file
        if self.file_path != path:
            self.reset_history()

        if path.lower().endswith(".xls"):
            reply = QMessageBox.question(
                self, "Convert to .xlsx?",
                ("The selected file is in the older .xls format. Convert now?"),
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            if reply == QMessageBox.StandardButton.Yes:
                try:
                    self.save_version("Original .xls Pre-Conversion Backup", target_path=path)
                    new_path = convert_xls_to_xlsx(path)
                    QMessageBox.information(self, "Success", f"Converted to: {os.path.basename(new_path)}")
                    self.file_path = new_path
                    self.set_file_label(self.file_path)
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Unable to convert file:\n{e}")
                    self.file_path = None
                    self.set_no_file_label()
                return
            else:
                self.file_path = None
                self.set_no_file_label()
                return
        else:
            self.file_path = path
            self.set_file_label(self.file_path)

    # UPDATE existing browse_file to use this new helper:
    def browse_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel files (*.xlsx *.xls)"
        )
        self.process_selected_file(path)

    def init_ui(self):
        root = QVBoxLayout(self)
        self.change_btn = QPushButton("Change File")
        self.remove_btn = QPushButton("Remove")
        self.open_file_btn = QPushButton("📄 Open File")
        self.open_folder_btn = QPushButton("📁 Open Folder")
        
        # ---- Header ----
        header = QHBoxLayout()
        
        # 1. Setup the Logo
        logo_label = QLabel()
        
        # Decode base64 and load into QPixmap
        pixmap = QPixmap()
        pixmap.loadFromData(QByteArray.fromBase64(logo_base64.encode('utf-8')))
        
        # Scale to 70x70 smoothly
        pixmap = pixmap.scaled(
            70, 70, 
            Qt.AspectRatioMode.KeepAspectRatio, 
            Qt.TransformationMode.SmoothTransformation
        )
        logo_label.setPixmap(pixmap)
        logo_label.setFixedSize(70, 70)

        def open_mcmaster_website(event):
            QDesktopServices.openUrl(QUrl("https://science.mcmaster.ca"))

        logo_label.mousePressEvent = open_mcmaster_website
        #logo_label.setCursor(Qt.CursorShape.PointingHandCursor) # Changes cursor on hover
        # -----------------------------------------
        
        # Add logo to the far left
        header.addWidget(logo_label)
        
        # Add a stretch to push the title to the center
        header.addStretch()
        
        # 2. Setup the Title and Subtitle together tightly
        title_layout = QVBoxLayout()
        title_layout.setSpacing(0) # This removes the gap between title and subtitle!
        
        title = QLabel("McMaster Research Group for Stable Isotopologues")
        title.setObjectName("titleLabel")
        title.setFont(QFont("Arial", 20, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # --- Add this line to make it act like a hidden href ---
        # Define a standard function so it returns None instead of a boolean
        def open_mrsi_website(event):
            QDesktopServices.openUrl(QUrl("https://mrsi.mcmaster.ca/"))

        title.mousePressEvent = open_mrsi_website

        # Optional: Make the mouse cursor turn into a pointing hand when hovering over it.
        # If you want it to be a 100% secret "Easter egg" link, just leave this next line out!
        #title.setCursor(Qt.CursorShape.PointingHandCursor)
        
        subtitle = QLabel("Data Normalization Tool")
        subtitle.setObjectName("subtitleLabel")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setFont(QFont("Arial", 14))
        
        # Add both to the mini vertical layout
        title_layout.addWidget(title)
        title_layout.addWidget(subtitle)
        
        # Add the combined text block to the header
        header.addLayout(title_layout)
        
        # Add another stretch to balance the center
        header.addStretch()
        
        # 3. Add a dummy widget to the right to keep the title perfectly centered
        dummy_label = QLabel()
        dummy_label.setFixedSize(70, 70) # MATCHES LOGO SIZE for perfect centering
        header.addWidget(dummy_label)
        
        root.addLayout(header)
        
        # ---- Lock Button ----
        #self.lock_btn = QPushButton("🔒", self)
        #self.lock_btn.setProperty("flat", True)
        #self.lock_btn.setFixedSize(36, 36)
        #self.lock_btn.setStyleSheet("""
        #            QPushButton {
        #                background-color: #3f51b5;
        #                color: white;
        #                border-radius: 18px;
        #                font-size: 18px;
        #                padding: 4px 0;
        #            }
        #            QPushButton:hover { background-color: #303f9f; }
        #        """)
        #self.lock_btn.clicked.connect(self.toggle_advanced_settings_lock)
        
        # ---- Menu Button ----
        self.menu_btn = LongPressButton("≡", self)
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
        self.menu_btn.longPressed.connect(self.trigger_secret_unlock)

        # ---- History Button (NEW) ----
        self.history_btn = QPushButton("↺", self)
        self.history_btn.setProperty("flat", True)
        self.history_btn.setFixedSize(36, 36)
        self.history_btn.setToolTip("Version History / Rollback")
        self.history_btn.setStyleSheet("""
            QPushButton {
                background-color: #FF9800;
                color: white;
                border-radius: 18px;
                font-size: 18px;
                padding: 4px 0;
            }
            QPushButton:hover { background-color: #F57C00; }
        """)
        self.history_btn.clicked.connect(self.show_history_dialog)
        self.history_btn.hide() # Hidden until a backup actually happens
        
        QTimer.singleShot(0, self.position_header_buttons)
        
        # ==========================================================
        # ==== UPDATED FILE SELECTION BOX (DRAG & DROP) ===========
        # ==========================================================
        self.file_box = DragDropBox("File Selection") # <--- Changed to custom class
        self.file_box.setObjectName("fileGroup")
        
        # Connect the drop signal
        self.file_box.fileDropped.connect(self.process_selected_file)

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
        
        # Add a subtle hint label next to the browse button
        hint_label = QLabel("Drag & Drop File Here or \u2192 ")
        hint_label.setStyleSheet("color: #888; font-style: italic; font-size: 12px;")
        
        self.open_file_btn.setFixedWidth(120)
        self.open_folder_btn.setFixedWidth(120)
        for b in [self.open_file_btn, self.open_folder_btn]:
            b.hide()
        self.open_file_btn.clicked.connect(self.open_file)
        self.open_folder_btn.clicked.connect(self.open_folder)
        
        btn_row.addWidget(hint_label) # Added hint
        btn_row.addWidget(self.browse_btn)
        btn_row.addWidget(self.open_file_btn)
        btn_row.addWidget(self.open_folder_btn)
        btn_row.addStretch()
        fl.addLayout(btn_row)
        self.file_box.setLayout(fl)
        root.addWidget(self.file_box)
        # ==========================================================
        
        # ---- Tabs and Content Frame ----
        tabs_and_content = QVBoxLayout()
        """
        self.tab_bar = QTabBar()
        self.tab_bar.addTab("Carbonate")
        self.tab_bar.addTab("Water")
        self.tab_bar.addTab("Combine Data")
        self.advanced_tab_index = -1
        """

        self.tab_bar = QTabBar()
        self.tab_bar.addTab("Water")
        self.tab_bar.addTab("Carbonate")
        self.combine_tab_index = -1  # <--- Added tracker
        self.advanced_tab_index = -1
        
        self.tab_bar.setExpanding(False)
        self.tab_bar.setMovable(False)
        self.tab_bar.setDrawBase(False)
        self.tab_bar.currentChanged.connect(self.on_tab_changed)
        
        tab_wrapper_layout = QHBoxLayout()
        tab_wrapper_layout.setContentsMargins(0, 0, 0, 0)
        
        tab_wrapper_layout.addStretch()         # Pushes from the left
        tab_wrapper_layout.addWidget(self.tab_bar)
        tab_wrapper_layout.addStretch()         # Pushes from the right
        
        tabs_and_content.addLayout(tab_wrapper_layout)
        
        self.content_frame = QFrame()
        self.content_frame.setObjectName("contentPane")
        self.content_frame.setFrameShape(QFrame.Shape.StyledPanel)
        
        self.content_frame_layout = QVBoxLayout(self.content_frame)
        self.content_frame_layout.setContentsMargins(8, 8, 8, 8)
        self.content_frame_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        tabs_and_content.addWidget(self.content_frame)
        root.addLayout(tabs_and_content)
        
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
        
        # ---- Run/Stop/Clear Button Container ----
        self.button_container = QHBoxLayout()
        self.run_btn = QPushButton("▶ Run Selected Steps")
        self.run_btn.clicked.connect(self.run_steps)
        
        self.stop_btn = QPushButton("◼ Stop")
        self.stop_btn.setObjectName("stopBtn")
        self.stop_btn.clicked.connect(self.stop_steps)
        self.stop_btn.hide()
        
        self.clear_log_btn = QPushButton("✖ Clear Log")
        self.clear_log_btn.setObjectName("clearBtn")
        self.clear_log_btn.clicked.connect(self.clear_log)
        
        self.button_container.addWidget(self.run_btn)
        self.button_container.addWidget(self.stop_btn)
        self.button_container.addWidget(self.clear_log_btn)
        root.addLayout(self.button_container)

        # Initialize tab state
        self.on_tab_changed(0)
    
    def clear_log(self):
        self.log_box.clear()
        self.log_box.hide()
        
        # Stop any running animation
        if self.progress_anim:
            self.progress_anim.stop()
            
        self.progress.setValue(0) # Reset instantly
        self.progress.hide()
        self.resize(self.width(), 700)
        self.run_btn.show()
        self.stop_btn.hide()

    def position_header_buttons(self):
        self.menu_btn.move(self.width() - 50, 10)
        self.history_btn.move(self.width() - 95, 10)
    
    # def update_lock_state(self):
    #     if self.is_locked:
    #         self.lock_btn.setText("🔒")
    #         self.lock_btn.setStyleSheet("""
    #             QPushButton { background-color: #808080; color: white; border-radius: 18px; font-size: 18px; padding: 4px 0; }
    #             QPushButton:hover { background-color: #808080; }
    #         """)
    #         if self.advanced_tab_index != -1:
    #             if self.tab_bar.currentIndex() == self.advanced_tab_index:
    #                 self.tab_bar.setCurrentIndex(0)
    #             self.tab_bar.removeTab(self.advanced_tab_index)
    #             self.advanced_tab_index = -1
    #     else:
    #         self.lock_btn.setText("🔓")
    #         self.lock_btn.setStyleSheet("""
    #             QPushButton { background-color: #FF9800; color: white; border-radius: 18px; font-size: 18px; padding: 4px 0; }
    #             QPushButton:hover { background-color: #F57C00; }
    #         """)
    #         if self.advanced_tab_index == -1:
    #             self.advanced_tab_index = self.tab_bar.addTab("Advanced Settings")

    def update_lock_state(self):
        if self.is_locked:
            self.lock_btn.setText("🔒")
            self.lock_btn.setStyleSheet("""
                QPushButton { background-color: #808080; color: white; border-radius: 18px; font-size: 18px; padding: 4px 0; }
                QPushButton:hover { background-color: #808080; }
            """)
            if self.tab_bar.currentIndex() in [self.advanced_tab_index, self.combine_tab_index]:
                self.tab_bar.setCurrentIndex(0)
            
            if self.advanced_tab_index != -1:
                self.tab_bar.removeTab(self.advanced_tab_index)
                self.advanced_tab_index = -1
            if self.combine_tab_index != -1:
                self.tab_bar.removeTab(self.combine_tab_index)
                self.combine_tab_index = -1
        else:
            self.lock_btn.setText("🔓")
            self.lock_btn.setStyleSheet("""
                QPushButton { background-color: #FF9800; color: white; border-radius: 18px; font-size: 18px; padding: 4px 0; }
                QPushButton:hover { background-color: #F57C00; }
            """)
            if self.combine_tab_index == -1:
                self.combine_tab_index = self.tab_bar.addTab("Combine Data")
            if self.advanced_tab_index == -1:
                self.advanced_tab_index = self.tab_bar.addTab("Advanced Settings")

    def toggle_advanced_settings_lock(self):
        if self.is_locked:
            popup = PasswordPopup(self)
            if popup.exec() == QDialog.DialogCode.Accepted:
                if popup.password_correct:
                    self.is_locked = False
                    self.update_lock_state()
                    if self.advanced_tab_index != -1:
                        self.tab_bar.setCurrentIndex(self.advanced_tab_index)
        else:
            self.is_locked = True
            self.update_lock_state()
    
    def show_toast(self, message):
        """Displays a temporary floating message bubble."""
        # Create the label without a layout (absolute positioning)
        toast = QLabel(message, self)
        toast.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # Style it to look like a notification bubble
        toast.setStyleSheet("""
            QLabel {
                background-color: #323232;
                color: white;
                padding: 12px 24px;
                border-radius: 20px;
                font-size: 14px;
                font-weight: bold;
                opacity: 200;
            }
        """)
        toast.adjustSize()
        
        # Position at bottom center
        x = (self.width() - toast.width()) // 2
        y = 40 #self.height() - #toast.height()
        toast.move(x, y)
        toast.show()
        
        # Raise to top ensure it sits over other widgets
        toast.raise_()
        
        # Auto-delete after 3000ms (3 seconds)
        QTimer.singleShot(3000, toast.deleteLater)
    
        """
        def trigger_secret_unlock(self):

            #Discrete toggle for settings.
            #If Locked -> Ask password -> Unlock.
            #If Unlocked -> Lock immediately (hide tab).

            # --- CASE 1: ALREADY UNLOCKED -> RE-LOCK IT ---
            if not self.is_locked:
                # 1. If we are currently looking at the Advanced tab, switch to the first tab
                if self.advanced_tab_index != -1:
                    if self.tab_bar.currentIndex() == self.advanced_tab_index:
                        self.tab_bar.setCurrentIndex(0) 
                    
                    # 2. Remove the tab
                    self.tab_bar.removeTab(self.advanced_tab_index)
                    self.advanced_tab_index = -1
                
                # 3. Set state to Locked and show bubble
                self.is_locked = True
                
                # --- UPDATED: Use Toast instead of Popup ---
                self.show_toast("🔐 Settings Re-Locked")
                return

            # --- CASE 2: LOCKED -> UNLOCK IT ---
            popup = PasswordPopup(self)
            if popup.exec() == QDialog.DialogCode.Accepted:
                if popup.password_correct:
                    self.is_locked = False
                    
                    # Create the tab if it doesn't exist
                    if self.advanced_tab_index == -1:
                        self.advanced_tab_index = self.tab_bar.addTab("Advanced Settings")
                    
                    # Switch to it immediately
                    self.tab_bar.setCurrentIndex(self.advanced_tab_index)
                    
                    # Optional: You can also use the bubble for unlocking if you prefer
                    # self.show_toast("🔓 Advanced Settings Unlocked")
                    self.show_toast("🔓 Advanced Settings have been unlocked.")
        """
    
    def trigger_secret_unlock(self):
        """
        Discrete toggle for settings.
        If Locked -> Ask password -> Unlock.
        If Unlocked -> Lock immediately (hide tab).
        """
        # --- CASE 1: ALREADY UNLOCKED -> RE-LOCK IT ---
        if not self.is_locked:
            # 1. If we are currently looking at a restricted tab, switch to the first tab
            if self.tab_bar.currentIndex() in [self.advanced_tab_index, self.combine_tab_index]:
                self.tab_bar.setCurrentIndex(0) 
                
            # 2. Remove the tabs (Remove highest index first to avoid shifting bugs)
            if self.advanced_tab_index != -1:
                self.tab_bar.removeTab(self.advanced_tab_index)
                self.advanced_tab_index = -1
            if self.combine_tab_index != -1:
                self.tab_bar.removeTab(self.combine_tab_index)
                self.combine_tab_index = -1
            
            # 3. Set state to Locked and show bubble
            self.is_locked = True
            
            self.show_toast("🔐 Settings Re-Locked")
            return

        # --- CASE 2: LOCKED -> UNLOCK IT ---
        popup = PasswordPopup(self)
        if popup.exec() == QDialog.DialogCode.Accepted:
            if popup.password_correct:
                self.is_locked = False
                
                # Create the tabs if they don't exist
                if self.combine_tab_index == -1:
                    self.combine_tab_index = self.tab_bar.addTab("Combine Data")
                if self.advanced_tab_index == -1:
                    self.advanced_tab_index = self.tab_bar.addTab("Advanced Settings")
                
                # Switch to it immediately
                self.tab_bar.setCurrentIndex(self.advanced_tab_index)
                
                self.show_toast("🔓 Advanced Settings have been unlocked.")
                
    def on_tab_changed(self, index):
        # 1. Clean Content Frame
        while self.content_frame_layout.count():
            item = self.content_frame_layout.takeAt(0)
            if item.widget():
                item.widget().hide()

        # 2. Handle Tab Selection and Visibility
        if index == 0:
            self.active_tab_widget = self.water_tab # Assuming you flipped them!
            if hasattr(self.water_tab, 'refresh_step_labels'):
                self.water_tab.refresh_step_labels()
            self.file_box.show()
        elif index == 1:
            self.active_tab_widget = self.carbonate_tab
            if hasattr(self.carbonate_tab, 'refresh_step_labels'):
                self.carbonate_tab.refresh_step_labels()
            self.file_box.show()
        elif index == self.combine_tab_index and not self.is_locked:
            self.active_tab_widget = self.combine_tab
            self.file_box.hide()
        elif index == self.advanced_tab_index and not self.is_locked:
            self.active_tab_widget = self.advanced_settings_tab
            self.file_box.hide()
        else:
            self.active_tab_widget = self.water_tab
            self.tab_bar.setCurrentIndex(0)
            self.file_box.show()

        # 3. Add Widget
        self.content_frame_layout.addWidget(self.active_tab_widget)
        self.active_tab_widget.show()

        # 4. Handle Button Visibility 
        # --- FIXED: Hide buttons on Advanced Settings Tab ---
        if self.active_tab_widget == self.advanced_settings_tab:
            self.run_btn.hide()
            self.stop_btn.hide()
            self.clear_log_btn.hide()
        else:
            # Normal logic for Water, Carbonate, and Combine tabs
            if not (self.thread and self.thread.isRunning()):
                self.run_btn.show()
                self.clear_log_btn.show()
                self.stop_btn.hide()
            else:
                self.run_btn.hide()
                self.stop_btn.show()
                self.clear_log_btn.show()
        
    def show_menu(self):
        # Pass 'self' so the menu inherits the current stylesheet!
        menu = QMenu(self) 
        
        about_action = menu.addAction("About")
        about_action.triggered.connect(self.show_about)
        
        # --- ADD THESE TWO LINES ---
        update_action = menu.addAction("Check for Updates")
        update_action.triggered.connect(self.check_for_updates)
        # ---------------------------
        
        self.toggle_dm_action = menu.addAction("Toggle Dark Mode")
        self.toggle_dm_action.triggered.connect(self.toggle_dark_mode)
        
        # Calculate position relative to the BUTTON
        pos = self.menu_btn.mapToGlobal(QPoint(0, self.menu_btn.height()))
        
        menu.exec(pos)

    def show_about(self):
        """Displays the custom About dialog."""
        about_dialog = AboutDialog(self)
        about_dialog.exec()

    def toggle_dark_mode(self):
        self.dark_mode = not self.dark_mode
        if self.dark_mode:
            self.setStyleSheet(self.dark_stylesheet)
            self.carbonate_tab.setStyleSheet(self.dark_stylesheet)
            self.water_tab.setStyleSheet(self.dark_stylesheet)
            self.combine_tab.setStyleSheet(self.dark_stylesheet)
            self.advanced_settings_tab.setStyleSheet(self.dark_stylesheet)
        else:
            self.setStyleSheet(self.light_stylesheet)
            self.carbonate_tab.setStyleSheet(self.light_stylesheet)
            self.water_tab.setStyleSheet(self.light_stylesheet)
            self.combine_tab.setStyleSheet(self.light_stylesheet)
            self.advanced_settings_tab.setStyleSheet(self.light_stylesheet)
    
    # ... (File handling methods) ...
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
        self.reset_history()
        self.set_no_file_label()

    def browse_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel files (*.xlsx *.xls)"
        )
        if not path:
            return

        if path.lower().endswith(".xls"):
            reply = QMessageBox.question(
                self, "Convert to .xlsx?",
                ("The selected file is in the older .xls format. Convert now?"),
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            if reply == QMessageBox.StandardButton.Yes:
                try:
                    new_path = convert_xls_to_xlsx(path)
                    QMessageBox.information(self, "Success", f"Converted to: {os.path.basename(new_path)}")
                    self.file_path = new_path
                    self.set_file_label(self.file_path)
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Unable to convert file:\n{e}")
                    self.file_path = None
                    self.set_no_file_label()
                return
            else:
                self.file_path = None
                self.set_no_file_label()
                return
        else:
            self.file_path = path
            self.set_file_label(self.file_path)

    def open_file(self):
        if not self.file_path: return
        try:
            if sys.platform.startswith("darwin"): subprocess.run(["open", self.file_path])
            elif os.name == "nt": os.startfile(self.file_path)
            elif os.name == "posix": subprocess.run(["xdg-open", self.file_path])
        except Exception as e: QMessageBox.critical(self, "Error", f"Unable to open file:\n{e}")

    def open_folder(self):
        if not self.file_path: return
        folder = os.path.dirname(self.file_path)
        try:
            if sys.platform.startswith("darwin"): subprocess.run(["open", folder])
            elif os.name == "nt": os.startfile(folder)
            elif os.name == "posix": subprocess.run(["xdg-open", folder])
        except Exception as e: QMessageBox.critical(self, "Error", f"Unable to open folder:\n{e}")

    # ... (Logging/Progress methods) ...
    def on_log(self, msg, color):
        cmap = {"red": "#FF6B6B", "green": "#4CAF50", "white": "#E8EAED"}
        text_color = cmap.get(color, "#E8EAED") if self.dark_mode else cmap.get(color, "#202124")
        self.log_box.append(f'<span style="color:{text_color};">{msg}</span>')

    def on_progress(self, done, total, step_name):
        if total == 0: return
        
        # 1. Calculate the "Base" percentage (steps fully completed)
        # e.g., if 1 of 7 steps is done, base is 14%
        base_pct = int((done / total) * 100)
        
        # 2. Setup Animation
        if self.progress_anim is None:
            self.progress_anim = QPropertyAnimation(self.progress, b"value")
        
        # Stop any existing animation to re-purpose it
        self.progress_anim.stop()

        # 3. Handle "All Done" vs "Running Step"
        if step_name == "done" or done == total:
            # Case: Everything finished. Fast animation to 100%
            self.progress_anim.setDuration(500) # 0.5 seconds
            self.progress_anim.setStartValue(self.progress.value())
            self.progress_anim.setEndValue(100)
            self.progress_anim.setEasingCurve(QEasingCurve.Type.OutQuad)
            
            # Color: Green
            self.progress.setStyleSheet("QProgressBar::chunk { background-color: #4CAF50; }")
        
        else:
            # Case: A step is running. 
            # Trickle Effect: Animate slowly towards the *next* milestone 
            # but stop just short (e.g., 95% of the way to the next step).
            
            step_size = 100 / total
            
            # Target is the base + almost the whole next step
            target_trickle = int(base_pct + (step_size * 0.95))
            if target_trickle > 99: target_trickle = 99
            
            # Duration: Set this long (e.g., 10s) so it moves slowly. 
            # If the step finishes in 1s, the next call will interrupt this 
            # and snap the bar forward, which looks natural.
            self.progress_anim.setDuration(10000) 
            
            # Start from wherever the bar is currently (to be smooth)
            # But ensure we don't go backwards if the bar overshot
            start_val = max(self.progress.value(), base_pct)
            
            self.progress_anim.setStartValue(start_val)
            self.progress_anim.setEndValue(target_trickle)
            
            # Use OutCubic: It slows down as it gets closer to the target,
            # making it look like it's "hanging" waiting for the step to finish.
            self.progress_anim.setEasingCurve(QEasingCurve.Type.OutCubic)
            
            # Color: Blue (Running)
            self.progress.setStyleSheet("QProgressBar::chunk { background-color: #3f51b5; }")

        # 4. Start the logic
        self.progress_anim.start()

    def on_thread_done(self):
        self.on_log("Processing complete.", "green")
        self.run_btn.show()
        self.stop_btn.hide()
        self.run_btn.setEnabled(True)
        self.stop_btn.setEnabled(True)

    def on_thread_stopped(self):
        self.on_log("🛑 **Process stopped by user.**\n", "red")
        self.run_btn.show()
        self.stop_btn.hide()
        self.run_btn.setEnabled(True)
        self.stop_btn.setEnabled(True)
        self.progress.setStyleSheet("QProgressBar::chunk { background-color: #F44336; }")

    def run_steps(self):
        # --- COMBINE TAB LOGIC ---
        if self.active_tab_widget == self.combine_tab:
            files, output_path = self.combine_tab.get_run_parameters()
            
            if not files:
                QMessageBox.warning(self, "Error", "No files selected to combine.")
                return
            if not output_path:
                QMessageBox.warning(self, "Error", "Please select an output path.")
                return
            
            self.log_box.show()
            self.log_box.clear()
            self.on_log(f"Starting Combination of {len(files)} files...", "white")
            self.on_log(f"Output: {output_path}", "white")
            self.on_log("...Logic to be implemented in future update...", "white")
            self.on_log("Done (Placeholder).", "green")
            return
        
        # --- Advanced Settings Safeguard ---
        if self.active_tab_widget == self.advanced_settings_tab:
            QMessageBox.warning(self, "Action Not Allowed", 
                                "You cannot run the process from the Advanced Settings tab.\n\n"
                                "Please switch to the 'Carbonate' or 'Water' tab to run steps.")
            return

        # Validation
        if not self.file_path:
            QMessageBox.warning(self, "Error", "Please select a file first!")
            return
        if self.file_path.lower().endswith(".xls"):
            QMessageBox.warning(self, "Incompatible File", "Please convert .xls to .xlsx before running.")
            return
        
        try:
            steps, sheet_name, filter_opt = self.active_tab_widget.get_run_parameters()
        except AttributeError:
            QMessageBox.critical(self, "Internal Error", "Could not retrieve parameters.")
            return

        current_tab_index = self.tab_bar.currentIndex()
        tab_type = "water" if current_tab_index == 0 else "carbonate"

        if not steps or not any(steps.values()):
            QMessageBox.warning(self, "Error", "Please select at least one step to run.")
            return

        # UI State Change
        self.log_box.show()
        self.progress.show()
        self.resize(self.width(), max(self.height(), 950)) 
        self.log_box.clear()

        if self.progress_anim:
            self.progress_anim.stop()
        self.progress.setValue(0)
        
        self.run_btn.hide()
        self.stop_btn.show()
        self.stop_btn.setEnabled(True)
        self.progress.setStyleSheet("QProgressBar::chunk { background-color: #3f51b5; }")

        # Add Thread Log Connection
        self.save_version("Auto-backup before running steps") 

        # Start Thread
        self.thread = WorkerThread(self.file_path, steps, sheet_name, filter_opt, tab_type)
        self.thread.log.connect(self.on_log)
        self.thread.progress.connect(self.on_progress)
        self.thread.finished.connect(self.on_thread_done)
        self.thread.stopped_early.connect(self.on_thread_stopped)
        self.thread.start()

    def stop_steps(self):
        if self.thread and self.thread.isRunning():
            self.thread.stop()
            self.stop_btn.setEnabled(False)
            self.on_log("Stopping process... will complete the current step.", "red")


    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, "menu_btn"):
            self.position_header_buttons()

    # ==========================================
    # ==== AUTO-UPDATER INTEGRATION METHODS ====
    # ==========================================
    
    def check_for_updates(self):
        """Starts the background thread to check GitHub for a new release."""
        self.show_toast("Checking for updates...")
        
        # Initialize and start the check thread
        self.updater_thread = AutoUpdater(mode="check")
        self.updater_thread.check_finished.connect(self.on_update_check_finished)
        self.updater_thread.error_occurred.connect(self.on_update_error)
        self.updater_thread.start()

    def on_update_check_finished(self, has_update, latest_version, download_url):
        """Triggered when the version check is complete."""
        if has_update:
            # Call the shared custom dialog imported from updater.py
            dialog = UpdateAvailableDialog(CURRENT_VERSION, latest_version, self)
            
            # exec() blocks until the user clicks one of the buttons
            if dialog.exec() == QDialog.DialogCode.Accepted:
                self.start_update_download(download_url)
            else:
                # Optional: Let them know the update was skipped
                if hasattr(self, 'show_toast'):
                    self.show_toast("Update postponed.")
        else:
            # --- NEW CLEAN BANNER/TOAST ---
            msg = f"You are running the latest version ({CURRENT_VERSION})."
            if hasattr(self, 'show_toast'):
                self.show_toast(msg)
            else:
                # Fallback just in case show_toast isn't implemented
                self.show_temporary_banner(msg)

    def start_update_download(self, url):
        """Sets up the UI progress dialog and starts the download thread."""
        # Create a blocking progress dialog
        self.update_progress_dialog = QProgressDialog("Downloading update...", None, 0, 100, self)
        self.update_progress_dialog.setWindowTitle("Updating MRSI DNT")
        self.update_progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.update_progress_dialog.setCancelButton(None) # Prevent cancelling to avoid partial files
        self.update_progress_dialog.setAutoClose(True)
        self.update_progress_dialog.show()

        # Initialize and start the download thread
        self.updater_thread = AutoUpdater(mode="download", url=url)
        self.updater_thread.progress_updated.connect(self.update_progress_dialog.setValue)
        self.updater_thread.download_finished.connect(self.on_update_download_finished)
        self.updater_thread.error_occurred.connect(self.on_update_error)
        self.updater_thread.start()

    def on_update_download_finished(self, download_path):
        """Triggered when the download completes. Hands off to the OS script."""
        if download_path:
            self.show_toast("Update downloaded! Restarting application...")
            # Slight delay so the user can read the toast before the app closes
            QTimer.singleShot(1500, lambda: apply_update_and_restart(download_path))
        else:
            QMessageBox.warning(self, "Update Failed", "The download finished but the file could not be found.")

    def on_update_error(self, error_msg):
        """Handles network or OS errors gracefully."""
        # Close the progress dialog if it's open
        if hasattr(self, 'update_progress_dialog') and self.update_progress_dialog.isVisible():
            self.update_progress_dialog.close()
            
        QMessageBox.critical(self, "Updater Error", f"An error occurred during the update process:\n\n{error_msg}")

"""     
if __name__ == "__main__":
    app = QApplication(sys.argv)

    # 1. Show Splash Screen
    splash = StartupSplashScreen()
    splash.show()

    # 2. Simulate Loading (The "Minimum 1 Second" Logic)
    # We use a loop and processEvents() to keep the animation smooth 
    # while delaying the main window slightly.
    
    start_time = time.time()
    while time.time() - start_time < 1.5:  # Adjust 1.5 to make it longer/shorter
        app.processEvents() # Keeps the GUI responsive
        
        # Calculate fake progress based on time passed
        elapsed = time.time() - start_time
        progress = int((elapsed / 1.5) * 100) 
        splash.update_progress(progress)
        
        time.sleep(0.01) # Small sleep to prevent CPU hogging

    # 3. Load Main Window
    # (Real imports happen here or happened above, but the user sees the splash now)
    win = DataToolApp()
    
    # 4. Swap Windows
    splash.close()
    win.show()
    
    sys.exit(app.exec())
    """