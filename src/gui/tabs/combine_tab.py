import os
from PyQt6.QtGui import QPainter, QColor, QPen, QFont, QDesktopServices
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QFileDialog, QGroupBox, QLineEdit, QAbstractItemView, 
    QMessageBox, QRadioButton, QButtonGroup, QTableWidget, QTableWidgetItem, QHeaderView, QStyle, QCheckBox
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QUrl

# Import the new separated worker implementations
from gui.tabs.combine_processors.combine_water import WaterCombineWorker
from gui.tabs.combine_processors.combine_carbonate import CarbonateCombineWorker

class DragDropBox(QGroupBox):
    filesDropped = pyqtSignal(list) 

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
                    self.update() 
                    return
        event.ignore()

    def dragLeaveEvent(self, event):
        self.drag_active = False
        self.update()

    def dropEvent(self, event):
        self.drag_active = False
        self.update()
        
        valid_files = []
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(('.xlsx', '.xls')):
                valid_files.append(path)
                
        if valid_files:
            self.filesDropped.emit(valid_files)
            event.accept()

    def paintEvent(self, event):
        super().paintEvent(event)

        if self.drag_active:
            painter = QPainter(self)
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            
            rect = self.contentsRect()
            overlay_color = QColor("#E3F2FD") 
            overlay_color.setAlpha(200) 
            
            painter.setBrush(overlay_color)
            
            pen = QPen(QColor("#2196F3"))
            pen.setWidth(3)
            pen.setStyle(Qt.PenStyle.DashLine)
            painter.setPen(pen)
            
            painter.drawRoundedRect(rect.adjusted(5, 5, -5, -5), 10, 10)
            
            painter.setPen(QColor("#0D47A1"))
            font = QFont("Arial", 16, QFont.Weight.Bold)
            painter.setFont(font)
            painter.drawText(rect, Qt.AlignmentFlag.AlignCenter, "📂 Drop Excel File(s) Here")


class FileTableWidget(QTableWidget):
    def paintEvent(self, event):
        super().paintEvent(event)
        
        if self.rowCount() == 0:
            painter = QPainter(self.viewport())
            rect = self.viewport().rect()
            painter.setPen(QColor("#888888"))
            font = self.font()
            font.setPointSize(14)
            font.setItalic(True)
            font.setBold(True)
            painter.setFont(font)
            painter.drawText(rect, Qt.AlignmentFlag.AlignCenter, "📥 Drag & Drop File(s) Here")

            
class CombineTab(QWidget):
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)  
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10) 

        self.mode_warning_label = QLabel("Please select either water or carbonate")
        self.mode_warning_label.setStyleSheet("color: #d32f2f; font-size: 11px; font-weight: bold;")
        self.mode_warning_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.mode_warning_label)

        # --- 0. Mode Configuration ---
        mode_layout = QHBoxLayout()
        
        self.btn_water = QPushButton("Water")
        self.btn_carbonate = QPushButton("Carbonate")
        self.btn_water.setCheckable(True)
        self.btn_carbonate.setCheckable(True)
        
        self.btn_water.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_carbonate.setCursor(Qt.CursorShape.PointingHandCursor)
        
        toggle_style = """
            QPushButton {
                background-color: #f3f3f3;
                border: 1px solid #ccc;
                border-radius: 4px;
                padding: 4px 12px;
                color: #333;
                min-width: 80px; 
            }
            QPushButton:hover {
                background-color: #e5e5e5;
            }
            QPushButton:checked {
                font-weight: bold;
                border: 2px solid #333;
                background-color: #e5e5e5; 
            }
        """
        self.btn_water.setStyleSheet(toggle_style)
        self.btn_carbonate.setStyleSheet(toggle_style)
        
        self.btn_water.clicked.connect(self._on_mode_clicked)
        self.btn_carbonate.clicked.connect(self._on_mode_clicked)

        mode_layout.addStretch()
        mode_layout.addWidget(self.btn_water)
        mode_layout.addWidget(self.btn_carbonate)
        mode_layout.addStretch()
        layout.addLayout(mode_layout)

        # --- 1. File Handling ---
        copy_group = QGroupBox("File Handling")
        copy_layout = QHBoxLayout()
        
        self.handling_group = QButtonGroup(self)
        
        self.radio_temp_copy = QRadioButton("Process data on temp files")
        self.radio_modify_orig = QRadioButton("Process data on original files")
        self.radio_temp_copy.setChecked(True) 
        
        self.handling_group.addButton(self.radio_temp_copy)
        self.handling_group.addButton(self.radio_modify_orig)
        
        copy_layout.addWidget(self.radio_temp_copy)
        copy_layout.addWidget(self.radio_modify_orig)
        copy_layout.addStretch() 
        
        copy_group.setLayout(copy_layout)
        layout.addWidget(copy_group)

        # --- 2. File List Section ---
        list_group = DragDropBox("Raw Files to Combine")
        list_group.filesDropped.connect(self._add_files_to_table)
        
        list_layout = QVBoxLayout()
        list_layout.setSpacing(2) 
        
        top_controls = QHBoxLayout()
        
        self.browse_files_btn = QPushButton(" Browse Files")
        folder_icon = self.style().standardIcon(QStyle.StandardPixmap.SP_DirIcon)
        self.browse_files_btn.setIcon(folder_icon)
        self.browse_files_btn.clicked.connect(self.add_files)
        self.browse_files_btn.setStyleSheet("padding: 5px 15px;")
        
        self.clear_btn = QPushButton(" Clear All")
        trash_icon = self.style().standardIcon(QStyle.StandardPixmap.SP_TrashIcon)
        self.clear_btn.setIcon(trash_icon)
        self.clear_btn.clicked.connect(self.clear_all)
        self.clear_btn.setStyleSheet("padding: 5px 15px;")

        top_controls.addWidget(self.browse_files_btn)
        top_controls.addStretch() 
        top_controls.addWidget(self.clear_btn)
        list_layout.addLayout(top_controls)
        
        self.file_table = FileTableWidget(0, 3)
        self.file_table.setHorizontalHeaderLabels(["File Name", "Default Sheet Name", ""])
        
        self.file_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.file_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.file_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        self.file_table.setColumnWidth(2, 40) 
        
        self.file_table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection) 
        self.file_table.setAlternatingRowColors(True)
        list_layout.addWidget(self.file_table)
        
        self.footer_hint = QLabel("Drag & drop more files anywhere in this box...")
        self.footer_hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.footer_hint.setStyleSheet("color: #999999; font-style: italic; font-size: 11px;")
        self.footer_hint.setContentsMargins(0, 0, 0, 0) 
        self.footer_hint.hide()
        list_layout.addWidget(self.footer_hint)

        list_group.setLayout(list_layout)
        layout.addWidget(list_group)
        
        # --- 3. Output Configuration ---
        output_group = QGroupBox("Final Combined Output")
        output_layout = QVBoxLayout()
        
        output_layout.setContentsMargins(10, 8, 10, 8) 
        output_layout.setSpacing(5) 
        
        row_out = QHBoxLayout()
        self.output_path_input = QLineEdit()
        
        desktop_dir = os.path.join(os.path.expanduser("~"), "Desktop")
        default_out_path = os.path.join(desktop_dir, "Combined_Normalization_Data.xlsx")
        self.output_path_input.setText(default_out_path)
        
        self.browse_out_btn = QPushButton(" Browse Files")
        self.browse_out_btn.setFixedWidth(130) 
        
        folder_icon = self.style().standardIcon(QStyle.StandardPixmap.SP_DirIcon)
        self.browse_out_btn.setIcon(folder_icon)
        self.browse_out_btn.clicked.connect(self.browse_output)
        
        row_out.addWidget(QLabel("Output File:"))
        row_out.addWidget(self.output_path_input)
        row_out.addWidget(self.browse_out_btn)
        
        output_layout.addLayout(row_out)

        action_row = QHBoxLayout()
        
        self.open_checkbox = QCheckBox("Open file upon completion of processing")
        self.open_checkbox.setChecked(True) 
        self.open_checkbox.setCursor(Qt.CursorShape.PointingHandCursor)
        
        self.btn_open_file = QPushButton(" Open File")
        self.btn_open_file.setFixedWidth(130) 
        self.btn_open_file.setCursor(Qt.CursorShape.PointingHandCursor)
        
        open_icon = self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon)
        self.btn_open_file.setIcon(open_icon)
        self.btn_open_file.clicked.connect(self.open_combined_file)
        
        action_row.addWidget(self.open_checkbox)
        action_row.addStretch() 
        action_row.addWidget(self.btn_open_file)
        
        output_layout.addLayout(action_row)

        output_group.setLayout(output_layout)
        layout.addWidget(output_group)
        
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    
    def dropEvent(self, event):
        files = []
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.endswith(('.xlsx', '.xls')):
                files.append(file_path)
        self._add_files_to_table(files)

    def _add_files_to_table(self, files):
        existing_files = []
        for i in range(self.file_table.rowCount()):
            item = self.file_table.item(i, 0)
            if item:
                existing_files.append(item.data(Qt.ItemDataRole.UserRole))
                
        if self.btn_water.isChecked():
            default_sheet = "ExportGB1.wke"
        elif self.btn_carbonate.isChecked():
            default_sheet = "ExportGB2.wke"
        else:
            default_sheet = ""
        
        for f in files:
            if f not in existing_files:
                row = self.file_table.rowCount()
                self.file_table.insertRow(row)
                
                filename = os.path.basename(f)
                path_item = QTableWidgetItem(filename)
                path_item.setData(Qt.ItemDataRole.UserRole, f) 
                path_item.setFlags(path_item.flags() & ~Qt.ItemFlag.ItemIsEditable) 
                self.file_table.setItem(row, 0, path_item)
                
                sheet_item = QTableWidgetItem(default_sheet)
                self.file_table.setItem(row, 1, sheet_item)
                
                del_btn = QPushButton("−") 
                del_btn.setFixedSize(24, 24)
                del_btn.setToolTip("Remove this file")
                del_btn.setCursor(Qt.CursorShape.PointingHandCursor)
                del_btn.setStyleSheet("""
                    QPushButton { 
                        background-color: #ff4d4d; 
                        color: white; 
                        border: none; 
                        border-radius: 4px; 
                        font-weight: bold; 
                        font-size: 16px; 
                        padding: 0px; 
                    }
                    QPushButton:hover { 
                        background-color: #d32f2f; 
                    }
                """)
                del_btn.clicked.connect(lambda _, r=path_item: self._remove_specific_row(r))
                
                cell_widget = QWidget()
                cell_layout = QHBoxLayout(cell_widget)
                cell_layout.setContentsMargins(0, 0, 0, 0)
                cell_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
                cell_layout.addWidget(del_btn)
                
                self.file_table.setCellWidget(row, 2, cell_widget)
            
            self._update_footer_visibility()

    def _remove_specific_row(self, item):
        row = self.file_table.row(item)
        if row >= 0:
            self.file_table.removeRow(row)
        self._update_footer_visibility()

    def update_default_sheets_in_table(self):
        if self.btn_water.isChecked():
            default_sheet = "ExportGB1.wke"
        elif self.btn_carbonate.isChecked():
            default_sheet = "ExportGB2.wke"
        else:
            default_sheet = "" 

        for row in range(self.file_table.rowCount()):
            current_text = self.file_table.item(row, 1).text()
            if current_text in ["ExportGB1.wke", "ExportGB2.wke", ""]:
                self.file_table.item(row, 1).setText(default_sheet)

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Raw Excel Files", "", "Excel Files (*.xlsx *.xls)"
        )
        if files:
            self._add_files_to_table(files)

    def remove_selected(self):
        rows = sorted(set(index.row() for index in self.file_table.selectedIndexes()), reverse=True)
        for row in rows:
            self.file_table.removeRow(row)

    def clear_all(self):
        self.file_table.setRowCount(0)
        self._update_footer_visibility()

    def browse_output(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Combined File", "Combined_Normalization_Data.xlsx", "Excel Files (*.xlsx)"
        )
        if path:
            self.output_path_input.setText(path)

    def get_run_parameters(self):
        file_list = []
        for row in range(self.file_table.rowCount()):
            file_list.append({
                "path": self.file_table.item(row, 0).data(Qt.ItemDataRole.UserRole),
                "sheet": self.file_table.item(row, 1).text().strip()
            })
            
        return {
            "mode": "water" if self.btn_water.isChecked() else "carbonate",
            "protect_originals": self.radio_temp_copy.isChecked(),
            "file_list": file_list,
            "output_path": self.output_path_input.text().strip(),
            "open_on_complete": self.open_checkbox.isChecked() 
        }

    def _on_mode_clicked(self, checked):
        sender = self.sender()
        
        if not checked:
            sender.setChecked(True)
            return

        if sender == self.btn_water:
            self.btn_carbonate.setChecked(False)
        else:
            self.btn_water.setChecked(False)
            
        self.mode_warning_label.hide()
        self.update_default_sheets_in_table()

    def open_combined_file(self):
        path = self.output_path_input.text().strip()
        if not path or not os.path.exists(path):
            QMessageBox.warning(self, "File Not Found", "The combined file has not been created yet or the path is invalid.")
            return
        
        QDesktopServices.openUrl(QUrl.fromLocalFile(path))
    
    def _update_footer_visibility(self):
        if self.file_table.rowCount() == 0:
            self.footer_hint.hide()
        else:
            self.footer_hint.show()
    

# =========================================================================
# ALL-IN-ONE COMBINE WORKER PROXY
# =========================================================================
class CombineWorker(QThread):
    """
    A Proxy QThread that delegates the work to either the WaterCombineWorker
    or CarbonateCombineWorker, preventing the need to change code elsewhere in the app.
    """
    log = pyqtSignal(str, str)
    progress = pyqtSignal(int, int, str)
    finished = pyqtSignal()
    error = pyqtSignal(str)
    stopped_early = pyqtSignal()

    def __init__(self, params):
        super().__init__()
        self.params = params
        self._worker = None

    def stop(self):
        if self._worker:
            self._worker.stop()

    def run(self):
        mode = self.params.get("mode")
        
        # Instantiate the proper separated worker logic
        if mode == "water":
            self._worker = WaterCombineWorker(self.params)
        else:
            self._worker = CarbonateCombineWorker(self.params)

        # Route the separated worker's signals exactly as the UI expects
        self._worker.log.connect(self.log.emit)
        self._worker.progress.connect(self.progress.emit)
        self._worker.finished.connect(self.finished.emit)
        self._worker.error.connect(self.error.emit)
        self._worker.stopped_early.connect(self.stopped_early.emit)

        # Run the target logic within THIS thread context
        self._worker.run()