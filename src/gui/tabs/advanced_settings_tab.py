from PyQt6.QtGui import QFont, QIcon, QCursor, QPainter, QColor, QPen, QAction, QKeySequence, QPixmap, QImage, QDesktopServices, QDoubleValidator
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QFileDialog, QTabWidget, QTextEdit, QCheckBox,
    QLineEdit, QComboBox, QGroupBox, QMessageBox, QMenu, QProgressBar, QFrame,
    QSizePolicy, QSpacerItem, QGridLayout, QTabBar, QDialog, QScrollArea, QButtonGroup, 
    QRadioButton, QListWidget, QAbstractItemView, QTableWidget, QTableWidgetItem, QHeaderView, QLayout,
    QToolTip, QStyleOptionGroupBox, QProgressDialog, QLabel, QStyle, QDoubleSpinBox,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer, QPoint, QRect, QSize, QPropertyAnimation, QEasingCurve, QByteArray, QUrl

import utils.settings as settings
import gui.main_window as main_window

class AdvancedSettingsTab(QWidget):
    def __init__(self):
        super().__init__()
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(0, 0, 0, 0)  # <-- ADD THIS: Removes outer layout padding
        
        # Scroll Area
        self.scroll_area = QScrollArea()
        self.scroll_area.setFrameShape(QFrame.Shape.NoFrame)  # <-- ADD THIS: Removes the drawn border
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
        # --- NEW: Floating Reset Button (No Layout) ---
        # Parent it directly to 'self' so it floats over the ScrollArea
        self.btn_reset = QPushButton("Reset to Default", self) 
        self.btn_reset.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_reset.setStyleSheet("""
            QPushButton {
                background-color: #f3f3f3;
                border: 1px solid #ccc;
                border-radius: 4px;
                padding: 3px 5px;
                color: #333;
                font-weight: bold;
                margin-right: 10px;
            }
            QPushButton:hover {
                background-color: #e5e5e5;
            }
        """)
        self.btn_reset.clicked.connect(self._reset_to_default)
        
        # Explicitly show the button (since it's not in a layout)
        self.btn_reset.show()

        # Proceed with generating the rest of the UI
        self._create_general_config()
        self._create_outlier_settings()
        self._create_calc_logic_section()
        self._create_material_tabs()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        # Dynamically keep the floating button in the top right corner
        if hasattr(self, 'btn_reset') and self.btn_reset.isVisible():
            # Force the button to calculate its required width
            self.btn_reset.adjustSize()
            
            # 25px from the right edge, 15px from the top edge
            x = self.width() - self.btn_reset.width() - 25
            y = 15
            
            self.btn_reset.move(x, y)
            
            # Bring to front every time we resize to ensure it stays on top
            self.btn_reset.raise_()
        
    def _create_general_config(self):
        group = QGroupBox("1. Conditional Formatting for Excel")
        layout = QVBoxLayout() 
        group.setLayout(layout)
        
        # --- Helper for Info Icon ---
        def create_info_label(tooltip_text):
            lbl = main_window.InstantTooltipLabel("ⓘ") 
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

        # --- Visual Example Row ---
        visual_layout = QHBoxLayout()
        visual_layout.setContentsMargins(0, 0, 0, 0) # Add a little breathing room below it
        
        # 1. Normal Cell (White)
        self.lbl_visual_good = QLabel()
        self.lbl_visual_good.setFixedSize(45, 25)
        self.lbl_visual_good.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_visual_good.setStyleSheet("""
            background-color: #FFFFFF; 
            color: #000000; 
            border: 1px solid #D4D4D4; 
            font-family: 'Segoe UI', sans-serif;
            font-size: 11px;
        """)
        
        # 2. Arrow
        lbl_arrow = QLabel("➔")
        lbl_arrow.setStyleSheet("color: #888; font-size: 16px; font-weight: bold;")
        lbl_arrow.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # 3. Highlighted Cell (Excel Red)
        self.lbl_visual_bad = QLabel()
        self.lbl_visual_bad.setFixedSize(45, 25)
        self.lbl_visual_bad.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_visual_bad.setStyleSheet("""
            background-color: #FFC7CE; 
            color: #9C0006; 
            border: 1px solid #FFC7CE;
            font-family: 'Segoe UI', sans-serif;
            font-size: 11px;
            font-weight: bold;
        """)
        
        # Add to layout
        visual_layout.addWidget(QLabel("<small style='color: gray;'><i>Example:</i></small>"))
        visual_layout.addWidget(self.lbl_visual_good)
        visual_layout.addWidget(lbl_arrow)
        visual_layout.addWidget(self.lbl_visual_bad)
        visual_layout.addStretch()
        
        layout.addLayout(visual_layout)

        # --- NEW: Checkbox to Enable/Disable (Moved above threshold label) ---
        toggle_layout = QHBoxLayout()
        self.chk_stdev = QCheckBox()
        
        # Get setting and apply initial text
        is_enabled = settings.get_setting("STDEV_THRESHOLD_ENABLED")
        is_enabled_bool = is_enabled if is_enabled is not None else False # Changed fallback to False
        self.chk_stdev.setChecked(is_enabled_bool)
        self.chk_stdev.setText("Enabled" if is_enabled_bool else "Disabled")
        
        # NEW: Bold it if it's disabled
        self.chk_stdev.setStyleSheet("font-weight: normal;" if is_enabled_bool else "font-weight: bold;")
        
        self.chk_stdev.stateChanged.connect(self._on_stdev_toggled)
        
        toggle_layout.addWidget(self.chk_stdev)
        toggle_layout.addStretch()
        layout.addLayout(toggle_layout)

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

        # --- Use QLineEdit for "Unlimited" decimals ---
        self.input_stdev = QLineEdit()
        self.input_stdev.setFixedWidth(55) 
        
        # NEW: Stylesheet to visually dim the input box when disabled
        self.input_stdev.setStyleSheet("""
            QLineEdit:disabled {
                background-color: #EAEAEA;
                color: #B0B0B0;
                border: 1px solid #D3D3D3;
            }
        """)
        
        validator = QDoubleValidator(0.0, 100.0, 99, self)
        validator.setNotation(QDoubleValidator.Notation.StandardNotation)
        self.input_stdev.setValidator(validator)
        
        current_stdev = float(settings.get_setting("STDEV_THRESHOLD") or 0.8)
        self.input_stdev.setText(f"{current_stdev:g}")
        
        self.input_stdev.editingFinished.connect(self._on_stdev_changed)
        self.input_stdev.textChanged.connect(self._on_text_changed_for_visual)
        
        row1.addWidget(self.input_stdev)
        
        # --- Create Custom Up/Down Buttons ---
        self.btn_up = QPushButton("▲")
        self.btn_down = QPushButton("▼")
        
        for btn in [self.btn_up, self.btn_down]:
            btn.setFixedSize(20, 13)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            # NEW: Added disabled styling for the buttons so they also look inactive
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
                QPushButton:disabled {
                    border: 1px solid #D3D3D3;
                    color: #D3D3D3;
                }
            """)
            
        self.btn_up.clicked.connect(self._step_up_stdev)
        self.btn_down.clicked.connect(self._step_down_stdev)
        
        spin_btn_layout = QVBoxLayout()
        spin_btn_layout.setSpacing(2)
        spin_btn_layout.setContentsMargins(0, 0, 0, 0)
        spin_btn_layout.addWidget(self.btn_up)
        spin_btn_layout.addWidget(self.btn_down)
        
        row1.addLayout(spin_btn_layout)
        row1.addStretch()
        
        # Set initial visual state (Renamed method to reflect we change state, not visibility)
        self._update_stdev_state(self.chk_stdev.isChecked())
        
        layout.addLayout(row1)
        self.layout.addWidget(group)
        
        self._update_visual_example(current_stdev)

    def _update_visual_example(self, current_limit):
        self.lbl_visual_good.setText(f"{current_limit:.3f}")
        self.lbl_visual_bad.setText(f"{current_limit:.3f}")

    # --- UPDATED: State handling methods ---
    def _on_stdev_toggled(self):
        is_enabled = self.chk_stdev.isChecked()
        
        # Toggle the text between Enabled and Disabled
        self.chk_stdev.setText("Enabled" if is_enabled else "Disabled")
        
        # NEW: Update the styling so "Disabled" is bold
        self.chk_stdev.setStyleSheet("font-weight: normal;" if is_enabled else "font-weight: bold;")
        
        settings.set_setting("STDEV_THRESHOLD_ENABLED", is_enabled)
        self._update_stdev_state(is_enabled)
        
    def _update_stdev_state(self, is_enabled):
        # Instead of .setVisible(), we now use .setEnabled() to dim them
        self.input_stdev.setEnabled(is_enabled)
        self.btn_up.setEnabled(is_enabled)
        self.btn_down.setEnabled(is_enabled)


    def _create_outlier_settings(self):
        group = QGroupBox("2. Outlier Settings")
        layout = QVBoxLayout() 
        group.setLayout(layout)
        
        # --- NEW: Visual Example Row for Outliers ---
        example_layout = QHBoxLayout()
        example_layout.setContentsMargins(0, 0, 0, 5) # Keeps it tucked up near the title
        
        # 1. The 'Example:' text
        lbl_example_text = QLabel("<small style='color: gray;'><i>Example:</i></small>")
        
        # 2. The Excel-style Cell
        lbl_outlier_cell = QLabel("<s>4.020</s>")
        lbl_outlier_cell.setFixedSize(45, 22) # Made slightly shorter to match a standard row height
        lbl_outlier_cell.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_outlier_cell.setStyleSheet("""
            background-color: #FFFFFF; 
            color: #FF0000; 
            border: 1px solid #D4D4D4; 
            font-family: 'Segoe UI', sans-serif;
            font-size: 11px;
        """)
        
        # Add them together and push them to the left
        example_layout.addWidget(lbl_example_text)
        example_layout.addWidget(lbl_outlier_cell)
        example_layout.addStretch()

        layout.addLayout(example_layout)
        
        # --- Helper for Info Icon ---
        def create_info_label(tooltip_text):
            lbl = main_window.InstantTooltipLabel("ⓘ") 
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
        self.rb_2sigma.setStyleSheet("font-weight: bold;")
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
        self.rb_excl_ind.setStyleSheet("font-weight: bold;") 
        
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
        group = QGroupBox("3. Data Selection")
        layout = QVBoxLayout()
        group.setLayout(layout)
        
        # --- Helper for Info Icon ---
        def create_info_label(tooltip_text):
            lbl = main_window.InstantTooltipLabel("ⓘ") 
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
        self.rb_s3_last6.setStyleSheet("font-weight: bold;") # NEW: Bold default
        self.rb_s3_last6_excl = QRadioButton("Last 6 Outliers Excluded (See Section 2)")
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
        self.rb_s7_all.setStyleSheet("font-weight: bold;") # NEW: Bold default
        self.rb_s7_outlier = QRadioButton("Outliers Excluded (See Section 2)")
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
        
        self.water_widget = main_window.MaterialTypeWidget("Water",
                                               ["Water Standards", "Col D", "Col E", "Col F (δ²H)", "Col G (δ¹⁸O SMOW)", "Col H", "Color"])
        self.tabs.addTab(self.water_widget, "Water")

        self.carb_widget = main_window.MaterialTypeWidget("Carbonate", 
                                              ["Col C (Name)", "Col D", "Col E", "Col F (d13C)", "Col G (d18O)", "Col H", "Color"])
        self.tabs.addTab(self.carb_widget, "Carbonate")
        
        self.layout.addWidget(self.tabs)

    def _step_up_stdev(self):
        """Increases the line edit value by 0.01"""
        try: val = float(self.input_stdev.text() or 0.0)
        except ValueError: val = 0.0
        
        new_val = val + 0.01
        self.input_stdev.setText(f"{new_val:g}")
        self._on_stdev_changed()

    def _step_down_stdev(self):
        """Decreases the line edit value by 0.01"""
        try: val = float(self.input_stdev.text() or 0.0)
        except ValueError: val = 0.0
        
        # Prevent it from going below 0
        new_val = max(0.0, val - 0.01)
        self.input_stdev.setText(f"{new_val:g}")
        self._on_stdev_changed()

    def _on_text_changed_for_visual(self, text):
        """Updates the visual example in real time as the user types"""
        try: val = float(text) if text else 0.0
        except ValueError: val = 0.0
        self._update_visual_example(val)

    def _on_stdev_changed(self):
        """Saves the QLineEdit text as a float into settings"""
        try: val = float(self.input_stdev.text() or 0.0)
        except ValueError: val = 0.0
        
        settings.set_setting("STDEV_THRESHOLD", val)

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

    def _reset_to_default(self):
        """Resets all UI elements to their default states."""
        # 1. Conditional Formatting defaults
        self.chk_stdev.setChecked(False) # Triggers _on_stdev_toggled
        self.input_stdev.setText("0.08") # Fixed to match settings.py exactly
        self._on_stdev_changed()         # Manually save the 0.08 to settings
        
        # 2. Outlier Settings defaults
        self.rb_2sigma.setChecked(True)
        self.rb_excl_ind.setChecked(True)
        
        # 3. Data Selection defaults
        self.rb_s3_last6.setChecked(True)
        self.rb_s7_all.setChecked(True)
        
        # 4. Table and Slope Group defaults (Carbonate)
        default_carb_mats = [
            {"col_c": "IAEA 603", "col_d": "", "col_e": "", "col_f": "2.46", "col_g": "-2.37", "col_h": "", "color": "green"},
            {"col_c": "LSVEC",    "col_d": "", "col_e": "-46.6", "col_f": "", "col_g": "", "col_h": "-26.7", "color": "lightblue"},
            {"col_c": "NBS 18",   "col_d": "", "col_e": "", "col_f": "-5.01", "col_g": "-23.01", "col_h": "", "color": "red"},
            {"col_c": "NBS 19",   "col_d": "", "col_e": "", "col_f": "1.95",  "col_g": "-2.20",  "col_h": "", "color": "darkblue"}
        ]
        default_carb_slopes = [
            ["NBS 18", "NBS 19"],
            ["NBS 18", "NBS 19", "IAEA 603"]
        ]
        settings.set_setting("REFERENCE_MATERIALS", default_carb_mats, sub_key="Carbonate")
        settings.set_setting("SLOPE_INTERCEPT_GROUPS", default_carb_slopes, sub_key="Carbonate")

        # 5. Table and Slope Group defaults (Water)
        default_water_mats = [
            {"col_c": "MRSI-STD-W1", "col_d": "", "col_e": "", "col_f": "-3.52", "col_g": "-0.58", "col_h": "", "color": "red"},
            {"col_c": "MRSI-STD-W2",  "col_d": "", "col_e": "", "col_f": "-214.79", "col_g": "-28.08", "col_h": "", "color": "darkblue"},
            {"col_c": "USGS W-67400",  "col_d": "", "col_e": "", "col_f": "1.25", "col_g": "-1.97", "col_h": "", "color": "orange"},
            {"col_c": "USGS W-64444",  "col_d": "", "col_e": "", "col_f": "-399.1", "col_g": "-51.14", "col_h": "", "color": "green"}
        ]
        default_water_slopes = [
            ["MRSI-STD-W1", "MRSI-STD-W2"],
            ["USGS W-67400", "USGS W-64444"]
        ]
        settings.set_setting("REFERENCE_MATERIALS", default_water_mats, sub_key="Water")
        settings.set_setting("SLOPE_INTERCEPT_GROUPS", default_water_slopes, sub_key="Water")

        # 6. Force the UI to refresh the tables and slope groups
        self.carb_widget.load_data()
        self.water_widget.load_data()

    """ old implementation without reseting table and slope intercept groups
    def _reset_to_default(self):
        # 1. Conditional Formatting defaults
        self.chk_stdev.setChecked(False) # Triggers _on_stdev_toggled
        self.input_stdev.setText("0.08")
        self._on_stdev_changed()         # Manually save the 0.8 to settings
        
        # 2. Outlier Settings defaults
        self.rb_2sigma.setChecked(True)
        self.rb_excl_ind.setChecked(True)
        
        # 3. Data Selection defaults
        self.rb_s3_last6.setChecked(True)
        self.rb_s7_all.setChecked(True)
    """