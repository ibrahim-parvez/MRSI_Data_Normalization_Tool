from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QCheckBox, QLineEdit, 
    QComboBox, QGroupBox, QFrame, QLabel
)
from PyQt6.QtCore import Qt

class CarbonateTab(QWidget):
    def __init__(self):
        super().__init__()
        self.step_cbs = []
        self.select_all_cb = None
        self.sheet_entry = None
        self.filter_combo = None
        self.step_names = [
            "Step 1: Data",
            "Step 2: To Sort",
            "Step 3: Last 6",
            "Step 4: Pre-Group",
            "Step 5: Group",
            "Step 6: Normalization",
            "Step 7: Report"
        ]
        self.layout = QVBoxLayout(self) 
        
        # This alignment setting is optional but recommended to keep content at the top.
        self.layout.setAlignment(Qt.AlignmentFlag.AlignTop) 
        
        self._create_step_ui()
        
    def _add_divider(self, layout):
        divider = QFrame()
        divider.setFrameShape(QFrame.Shape.HLine)
        divider.setFrameShadow(QFrame.Shadow.Sunken)
        # NOTE: Stylesheet needs to be managed by the parent or a custom property
        # For simplicity, we'll keep a minimal style here, assuming parent style is inherited
        # divider.setStyleSheet("color: #CCC;") 
        layout.addWidget(divider)

    def _create_step_ui(self):
        """Generates the steps UI and connects signals."""
        #layout = QVBoxLayout(self)
        
        # Steps group
        steps_group = QGroupBox("Carbonate Steps")
        steps_layout = QVBoxLayout()
        steps_group.setLayout(steps_layout)

        # Top: Select all toggle
        top_row = QHBoxLayout()
        top_row.addSpacing(10)
        self.select_all_cb = QCheckBox("Select All")
        top_row.addWidget(self.select_all_cb)
        top_row.addStretch()
        steps_layout.addLayout(top_row)
        steps_layout.addSpacing(8)
        
        for i, step_name in enumerate(self.step_names):
            row = QHBoxLayout()
            cb = QCheckBox(step_name)
            self.step_cbs.append(cb)
            
            row.addWidget(cb)
            
            # Special logic for Step 1 (Sheet Name)
            if step_name.startswith("Step 1"):
                self.sheet_entry = QLineEdit("Default_Gas_Bench.wke")
                self.sheet_entry.setFixedWidth(220)
                self.sheet_entry.setPlaceholderText("Sheet name")
                row.addSpacing(10)
                row.addWidget(QLabel("Sheet:"))
                row.addWidget(self.sheet_entry)

            # Special logic for Step 2 Carbonate (Filters)
            if step_name == "Step 2: To Sort":
                self.filter_combo = QComboBox()
                self.filter_combo.addItems(["All", "Last 6", "Ref Avg", "Start", "End", "Delta"])
                self.filter_combo.setCurrentText("Last 6")
                self.filter_combo.setFixedWidth(160)
                row.addSpacing(10)
                row.addWidget(QLabel("Filter:"))
                row.addWidget(self.filter_combo)
            
            row.addStretch()
            steps_layout.addLayout(row)
            
            if i < len(self.step_names) - 1:
                self._add_divider(steps_layout)

        # Add the main group to the tab's main layout (self.layout)
        self.layout.addWidget(steps_group) 
        
        # Add the stretch to push the 'steps_group' to the top
        self.layout.addStretch()

        # --- Link signals ---
        self.select_all_cb.stateChanged.connect(self.toggle_select_all)
        for cb in self.step_cbs:
            cb.stateChanged.connect(self.update_select_all_state)
        self.update_select_all_state() # Initial state check

    # --- UI Logic Methods ---
    def toggle_select_all(self, state):
        """Toggle all step checkboxes to match the Select All checkbox."""
        is_checked = self.select_all_cb.isChecked()
        for cb in self.step_cbs:
            cb.blockSignals(True)
            cb.setChecked(is_checked)
            cb.blockSignals(False)

    def update_select_all_state(self):
        """Keep Select All checkbox in sync with individual checkboxes."""
        all_checked = all(cb.isChecked() for cb in self.step_cbs) if self.step_cbs else False
        
        self.select_all_cb.blockSignals(True)
        self.select_all_cb.setChecked(all_checked)
        self.select_all_cb.blockSignals(False)
        
    def get_run_parameters(self):
        """Returns the steps dictionary, sheet name, and filter option."""
        steps = {name: cb.isChecked() for name, cb in zip(self.step_names, self.step_cbs)}
        sheet_name = self.sheet_entry.text().strip() if self.sheet_entry else "Default_Sheet"
        filter_opt = self.filter_combo.currentText() if self.filter_combo else "N/A"
        return steps, sheet_name, filter_opt