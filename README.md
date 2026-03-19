# McMaster Research Group for Stable Isotopologues Data Normalization Tool

This application is designed to fully automate and streamline the processing of mass spectrometer data for the McMaster Research Group for Stable Isotopologues. Built to be robust and user-friendly, this release introduces a complete graphical interface for handling both Carbonate and Water sample data with highly customizable statistical parameters.

* Download the Installer: https://github.com/ibrahim-parvez/MRSI_Data_Normalization_Tool/releases/tag/Installer

---

## ✨ Key Features

### Core Processing Workflows

* **Dedicated Sample Pipelines:** Isolated, fully-featured processing tabs for both Carbonate and Water isotopic data.
* **7-Step Granular Execution:** Run specific steps of the normalization process individually or all at once (Data Loading, Sorting, Last 6 Averaging, Pre-Grouping, Grouping, Normalization, and Final Reporting).
* **Combine Data:** A foundational interface for merging multiple processed datasets into a single output.

### Data & File Handling

* **Smart Drag-and-Drop:** Easily load datasets by dragging Excel files directly into the application window.
* **Legacy Format Support:** Automatic detection and conversion of older `.xls` files into `.xlsx` formats using `pandas` and `openpyxl` under the hood.
* **Live Excel Refreshing:** Uses `xlwings` to automatically recalculate and refresh Excel formulas during processing steps to ensure data integrity.

### Advanced Configuration & Analytics

* **Secured Advanced Settings:** A password-protected settings panel to prevent accidental modification of core mathematical logic.
* **Customizable Outlier Detection:** Configure strictness with 1σ, 2σ, or 3σ standard deviation thresholds for outlier calculation.
* **Dynamic Exclusion Logic:** Choose whether to discard entire rows when an outlier is found, or individually keep valid Carbon/Oxygen measurements.
* **Reference Material Management:** Fully editable tables to add, remove, and color-code reference materials and define custom Slope/Intercept calculation groups.

### User Experience (UX)

* **Modern GUI:** Built with PyQt6, featuring a custom splash screen and responsive design.
* **Dark & Light Mode:** Toggleable visual themes to suit user preference and reduce eye strain.
* **Live Execution Logging:** A real-time log box and interactive progress bar to track exactly which steps are running, completed, or failed.

---

## 🛠️ Technical Notes

* Packaged as a standalone executable (`.exe` for Windows / `.app` for macOS) requiring no prior Python installation for end-users.
* Includes dynamic resource pathing to ensure assets like application logos and splash screen graphics render correctly in distributed builds.
