import os
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QFileDialog, QListWidget, QGroupBox, 
    QLineEdit, QAbstractItemView, QMessageBox
)
from PyQt6.QtCore import Qt

class CombineTab(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # --- 1. File List Section ---
        list_group = QGroupBox("Files to Combine")
        list_layout = QVBoxLayout()
        
        self.file_list_widget = QListWidget()
        self.file_list_widget.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.file_list_widget.setAlternatingRowColors(True)
        list_layout.addWidget(self.file_list_widget)

        # Buttons for list management
        btn_layout = QHBoxLayout()
        
        self.add_btn = QPushButton("➕ Add Files")
        self.add_btn.clicked.connect(self.add_files)
        
        self.remove_btn = QPushButton("➖ Remove Selected")
        self.remove_btn.clicked.connect(self.remove_selected)
        
        self.clear_btn = QPushButton("🗑 Clear All")
        self.clear_btn.clicked.connect(self.clear_all)

        btn_layout.addWidget(self.add_btn)
        btn_layout.addWidget(self.remove_btn)
        btn_layout.addWidget(self.clear_btn)
        list_layout.addLayout(btn_layout)
        
        list_group.setLayout(list_layout)
        layout.addWidget(list_group)

        # --- 2. Output Configuration ---
        output_group = QGroupBox("Output Settings")
        output_layout = QVBoxLayout()
        
        row_out = QHBoxLayout()
        self.output_path_input = QLineEdit()
        self.output_path_input.setPlaceholderText("Select where to save the combined file...")
        self.browse_out_btn = QPushButton("Browse")
        self.browse_out_btn.clicked.connect(self.browse_output)
        
        row_out.addWidget(QLabel("Output File:"))
        row_out.addWidget(self.output_path_input)
        row_out.addWidget(self.browse_out_btn)
        
        output_layout.addLayout(row_out)
        output_group.setLayout(output_layout)
        layout.addWidget(output_group)
        
        # Add stretch to push everything up
        layout.addStretch()

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Excel Files", "", "Excel Files (*.xlsx *.xls)"
        )
        if files:
            # Avoid duplicates
            existing_items = [self.file_list_widget.item(i).text() for i in range(self.file_list_widget.count())]
            for f in files:
                if f not in existing_items:
                    self.file_list_widget.addItem(f)

    def remove_selected(self):
        for item in self.file_list_widget.selectedItems():
            self.file_list_widget.takeItem(self.file_list_widget.row(item))

    def clear_all(self):
        self.file_list_widget.clear()

    def browse_output(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Combined File", "combined_data.xlsx", "Excel Files (*.xlsx)"
        )
        if path:
            self.output_path_input.setText(path)

    def get_run_parameters(self):
        """
        Returns the data needed for the main window to run the process.
        Returns: (file_list, output_path)
        """
        file_list = [self.file_list_widget.item(i).text() for i in range(self.file_list_widget.count())]
        output_path = self.output_path_input.text().strip()
        return file_list, output_path