from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QCheckBox, QLineEdit, 
    QComboBox, QGroupBox, QFrame, QLabel
)
from PyQt6.QtCore import Qt
import settings

class WaterTab(QWidget):
    def __init__(self):
        super().__init__()
        self.step_cbs = []
        self.select_all_cb = None
        self.sheet_entry = None
        self.filter_combo = None 
        self.step3_label = None  # Reference to the Step 3 label for updating
        
        # Base names used for logic keys (backend compatibility)
        self.base_step_names = [
            "Step 1: Data",
            "Step 2: To Sort",
            "Step 3: Last 6",
            "Step 4: Pre-Group",
            "Step 5: Group",
            "Step 6: Normalization",
            "Step 7: Report"
        ]
        
        self.layout = QVBoxLayout(self) 
        self.layout.setAlignment(Qt.AlignmentFlag.AlignTop) 
        
        self._create_step_ui()
        self.refresh_step_labels() # Initial check

    def _add_divider(self, layout):
        divider = QFrame()
        divider.setFrameShape(QFrame.Shape.HLine)
        divider.setFrameShadow(QFrame.Shadow.Sunken)
        layout.addWidget(divider)

    def _create_step_ui(self):
        """Generates the steps UI and connects signals."""
        steps_group = QGroupBox("Water Steps")
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
        
        for i, step_name in enumerate(self.base_step_names):
            row = QHBoxLayout()
            
            # --- Special Handling for Step 3 (Rich Text Support) ---
            if "Step 3" in step_name:
                # 1. Create an empty checkbox
                cb = QCheckBox() 
                self.step_cbs.append(cb)
                row.addWidget(cb)
                
                # 2. Create a QLabel for the text (so we can use HTML)
                self.step3_label = QLabel(step_name)
                row.addWidget(self.step3_label)
                
                # 3. Make the Label clickable (so it toggles the checkbox)
                def toggle_cb(event, checkbox=cb):
                    checkbox.setChecked(not checkbox.isChecked())
                self.step3_label.mousePressEvent = toggle_cb
                
            else:
                # Standard Checkbox for everything else
                cb = QCheckBox(step_name)
                self.step_cbs.append(cb)
                row.addWidget(cb)
            
            # --- Extra Widgets (Step 1 & Step 2) ---
            # Special logic for Step 1 (Sheet Name)
            if step_name.startswith("Step 1"):
                self.sheet_entry = QLineEdit("ExportGB1.wke") 
                self.sheet_entry.setFixedWidth(220)
                self.sheet_entry.setPlaceholderText("Sheet name")
                row.addSpacing(10)
                row.addWidget(QLabel("Sheet:"))
                row.addWidget(self.sheet_entry)

            # Special logic for Step 2 Water (Filters)
            if step_name == "Step 2: To Sort":
                self.filter_combo = QComboBox()
                self.filter_combo.addItems(["Last 6", "Last 6 Outliers Excl.", "all", "Amp 44", "delta", "end 11", "ref avg", "sparkline", "start 6"]) 
                self.filter_combo.setCurrentText("Last 6")
                self.filter_combo.setFixedWidth(160)
                row.addSpacing(10)
                row.addWidget(QLabel("Filter:"))
                row.addWidget(self.filter_combo)
            
            row.addStretch()
            steps_layout.addLayout(row)
            
            if i < len(self.base_step_names) - 1:
                self._add_divider(steps_layout)

        self.layout.addWidget(steps_group)
        self.layout.addStretch() # Push everything to the top

        # --- Link signals ---
        self.select_all_cb.stateChanged.connect(self.toggle_select_all)
        for cb in self.step_cbs:
            cb.stateChanged.connect(self.update_select_all_state)
        self.update_select_all_state() 
    
    def refresh_step_labels(self):
        """
        Dynamically updates the Step 3 label with HTML if needed.
        """
        calc_mode_step3 = settings.get_setting("CALC_MODE_STEP3")
        
        if self.step3_label:
            if calc_mode_step3 == "Last 6 Outliers Excl.":
                # QLabel supports HTML rendering
                self.step3_label.setText("Step 3: Last 6 <span style='color:red;'><b>Outliers Excluded</b></span>")
            else:
                self.step3_label.setText("Step 3: Last 6")

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
        # Use base_step_names to ensure keys are consistent regardless of label HTML
        steps = {name: cb.isChecked() for name, cb in zip(self.base_step_names, self.step_cbs)}
        sheet_name = self.sheet_entry.text().strip() if self.sheet_entry else "Default_Sheet"
        filter_opt = self.filter_combo.currentText() if self.filter_combo else "N/A"
        return steps, sheet_name, filter_opt