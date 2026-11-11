<div align="center">
  <img src="assets/images/mrsi_logo.png" alt="MRSI" width="140">

  # MRSI Data Normalization Tool

  **Isotope-ratio mass spectrometer data, normalized and reported — without touching a spreadsheet formula.**

  Built for the McMaster Research Group for Stable Isotopologues

  [![Latest release](https://img.shields.io/github/v/release/ibrahim-parvez/MRSI_Data_Normalization_Tool?label=latest&color=7A003C)](https://github.com/ibrahim-parvez/MRSI_Data_Normalization_Tool/releases/latest)
  [![Checks](https://github.com/ibrahim-parvez/MRSI_Data_Normalization_Tool/actions/workflows/checks.yml/badge.svg)](https://github.com/ibrahim-parvez/MRSI_Data_Normalization_Tool/actions/workflows/checks.yml)

  Windows & macOS (Apple Silicon) · Python 3.13 · PyQt6

  [**Download the installer →**](https://github.com/ibrahim-parvez/MRSI_Data_Normalization_Tool/releases/tag/Installer)
</div>

---

## What this is

Raw output from an IRMS run is a wall of numbers. Turning it into publishable δ¹³C and δ¹⁸O values means sorting runs into groups, averaging the right replicates, flagging outliers, fitting a slope and intercept against your reference materials, normalizing every sample against that fit, and laying the whole thing out so the next person can follow it. Done by hand in Excel it takes an afternoon per run, and every manual step is a chance to paste into the wrong column.

The DNT does all of it in one click, and writes the result back into a **`_DNT` copy of your workbook** as a set of new sheets — so the raw data, every intermediate stage, and the final report all live together and stay auditable, and the file you dropped in is never modified.

It handles **Carbonate** and **Water** in dedicated pipelines, with optional **CO₂** and **Yield** calculations layered on for carbonate work.

---

## Highlights

| | |
|---|---|
| **Two full pipelines** | Carbonate and Water, each with its own reference materials, normalization maths, and report layout |
| **Run all 7 steps or just one** | Tick exactly the steps you need — re-run only the normalization after changing a setting, for example |
| **Batch + Combine** | Point it at a folder's worth of raw files, process them all, and merge every standard into one comparison workbook with charts |
| **Outlier control** | 1σ / 2σ / 3σ thresholds, and a choice between discarding a whole run or keeping the good half of it |
| **Editable reference materials** | Add, remove, recolor and regroup your standards — no code changes |
| **Your originals stay safe** | Work happens on a `_DNT` copy, with automatic backups before every run and crash-safe saves |
| **Drag and drop or browse** | Drop an `.xlsx` (or a legacy `.xls`, converted automatically) on the window, or use Browse Files — either way the sheet and materials are read in for you |
| **Updates itself** | Checks for new versions on launch and installs them in one click |
| **Dark mode** | `Ctrl/Cmd + D` |

---

## The processing pipeline

Every step reads the sheet the previous one produced and appends a new `_DNT` sheet. Nothing is overwritten, so you can open any stage and see exactly what happened.

| Step | Creates | What it does |
|:---:|---|---|
| **1** | `Data_DNT` | Reads the raw export, finds the header row wherever it sits, groups replicate runs, computes per-group C/O averages and standard deviations, flags outliers |
| **2** | `To Sort_DNT` | Converts formulas to cached values and applies the run filter (Last 6, Ref Avg, Delta, …) |
| **3** | `Last 6_DNT` | Keeps the analytically meaningful replicates — optionally with statistical outliers already removed |
| **4** | `Pre-Group_DNT` | Stages rows for grouping and copies formatting across |
| **5** | `Group_DNT` | Sorts reference materials and samples into blocks, marks outliers with red strikethrough, writes per-block Average / Stdev / Count |
| **6** | `Normalization_DNT` | The core: fits slope and intercept from your standards, normalizes every sample to VPDB and VSMOW, builds the blue reference box, and adds the CO₂ and Yield tables when enabled |
| **7** | `Report_DNT` | The clean, presentable summary sheet |

After step 7 the tool re-runs step 6 as a **"Finalizing Formatting"** pass, so the workbook opens on `Normalization_DNT` with everything laid out.

### How runs are grouped

Reference materials (HeCO₂, NBS 18, NBS 19, Carrara, …) are always gathered into one group each, since every run of a standard belongs in its slope.

Samples are grouped by name **per measurement session**. Runs are taken in Time Code order, and a new group starts whenever the run number goes back down:

- `A r1, A r2, A r3` … later `A r1, A r2, A r3` again → **two groups** (e.g. DAY 12 and DAY 9 of a time series), kept next to each other.
- `X r1, X r2` … later `X r3` → **one group**: the numbering carries on, so it is the same sample measured in two sittings.

### Carbonate extras

- **CO₂ table** — a parallel normalization against CO₂ reference values, with a Quick Switch for the 25 °C and 72 °C acid-fractionation profiles, or a **Custom** profile you can name yourself (there's a `°` insert button so you can type `70 ° C` without hunting for the character).
- **Yield table** — theoretical vs. calculated yield per sample, driven by a compound you choose for reference materials and a second one for samples. The two are color-coded (green for reference, pink for sample) all the way through the Theoretical column so you can see at a glance which is which. Each Yield slope uses only the runs whose name *is* that reference material — a sample called `A` is never mistaken for `Carrara`.

---

## Combine Data

For comparing standards across many runs.

1. Choose **Water** or **Carbonate**, then drop in as many raw files as you like.
2. For each file, drag identifiers between the **References** and **Samples** lists — per-file overrides, so a material treated as a standard in one run can be a sample in another.
3. Choose whether to work on **temporary copies** (originals untouched) or **the original files**.
4. Choose **Create New Combined File** or **Append to an Existing** one.

Each file is taken through all seven steps, then every reference material gets its own sheet in the combined workbook, with each run's block stacked in time order and native Excel scatter charts plus matplotlib trend plots drawn underneath.

---

## Advanced Settings

Advanced Settings are password-protected; ask the lab manager for access. Once unlocked, two extra tabs appear: **Combine Data** and **Advanced Settings**.

**Outliers**
- **Sigma threshold** — 1σ, 2σ or 3σ. Anything outside mean ± (σ × stdev) is flagged.
- **Exclusion logic** — *Individual* keeps a run's good δ¹⁸O even when its δ¹³C is an outlier; *Exclude Entire Row* discards the whole measurement if either value fails.

**Calculation modes**
- **Step 3** — `Last 6` or `Last 6 Outliers Excluded`.
- **Step 7** — report from `All Values` or `Outliers Excluded`.

**Standard deviation threshold** — on by default; highlights any group whose stdev exceeds your limit.

**Reference materials** — one editable table per type (Carbonate / Water / CO₂): name, published δ values, and display color. A material dragged into the References list on the Carbonate or Water tab is added here automatically with a color of its own. **Slope & Intercept Groups** define which standards are fitted together, and you can have several groups side by side.

**Yield** — atomic weights and compound definitions, plus the reference/sample compound selection.

**Reset to Default** puts everything back to the values the app launches with. **Manage Settings** keeps a history of snapshots (taken when you leave the tab, reset, restore or import), shows exactly how each one differs from your current settings, and restores any of them. Every reset, restore and import can be undone from there.

> Settings live in memory for the session and reset to defaults on relaunch. Use **Export / Import** in Manage Settings to keep a configuration.

---

## Keyboard shortcuts

| Shortcut | Action |
|---|---|
| `Ctrl/Cmd + 1` / `2` | Water / Carbonate tab |
| `Ctrl/Cmd + R` | Run the selected steps |
| `Ctrl/Cmd + D` | Toggle dark mode |
| `Ctrl/Cmd + /` | Show the shortcuts (also in the ≡ menu) |

The **Run Selected Steps** button stays greyed out until a run can start; hover over it to see what is still missing.

---

## Your data is protected

Processing writes into a workbook, so the tool takes that seriously:

- **Your original is never touched.** Steps run on a `_DNT` copy next to the file you chose.
- **A backup before every run.** Snapshots are kept for the session and reachable from the History button, so you can view or restore any earlier state.
- **Crash-safe saves.** Every sheet is written to a temporary file and swapped into place atomically. If the disk fills, a network drive drops, or the machine dies mid-save, your workbook is either the old version or the new one — never a half-written file.
- **Read-only files are respected** rather than silently overwritten.
- **Temp-copy mode** in Combine leaves your raw files byte-for-byte untouched.
- **Steps fail loudly.** If a prerequisite sheet is missing, you get a message naming the sheet and the step to run first — never a green "completed" on a step that did nothing.

---

## Installing

### For end users

1. Download the installer for your computer from the [Installer release](https://github.com/ibrahim-parvez/MRSI_Data_Normalization_Tool/releases/tag/Installer): `MRSI_DNT_Installer.exe` on Windows, `MRSI_DNT_Installer.app.zip` on a Mac.
2. Open it, pick a version (the newest is selected) and a folder, and press **Download and Install**. No Python needed.

**Updates are automatic.** Each time the app starts it checks for a newer release. If there is one, it shows what changed and offers **Update Now**: the new version is downloaded, checked against the release's checksum, swapped in, and the app reopens. You can also check at any time from **≡ → Check for Updates**. If an update can't be installed (for example, the app sits in a folder you can't write to), your current version is left exactly as it was.

**Microsoft Excel must be installed.** The tool drives a hidden Excel instance to recalculate formulas between steps, which only Excel can do correctly.

The apps are not code-signed, so the first launch needs one extra click:
- **Windows:** if SmartScreen says *"Windows protected your PC"*, choose **More info → Run anyway**.
- **macOS:** if it says the app *can't be opened*, right-click it and choose **Open**, or allow it under **System Settings → Privacy & Security**. The Mac build is for Apple Silicon (M1 or newer).

On macOS you'll also be asked to allow the app to control Microsoft Excel the first time it recalculates — this is xlwings driving Excel via AppleScript, and it needs to be granted for processing to work.

### From source

```bash
git clone https://github.com/ibrahim-parvez/MRSI_Data_Normalization_Tool.git
cd MRSI_Data_Normalization_Tool

python3 --version                  # needs 3.13
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate

pip install -r requirements.txt

python src/main.py
```

Running from source skips the update check; updates only apply to the installed app.

---

## Releasing a new version

Releases are built by **GitHub Actions** on real Windows and macOS machines, so there's no need for a Windows VM or a local PyInstaller run.

1. Commit and push your changes to `main`. The **Checks** workflow starts the app on Windows and macOS within a few minutes; a green tick on the commit means it's safe to release.
2. On GitHub, open **Actions → Release → Run workflow**.
3. Choose the bump — **patch** for fixes (2.1.1 → 2.1.2), **minor** for new features (2.1.1 → 2.2.0), **major** for big changes (2.1.1 → 3.0.0).
4. Type what changed, separating items with semicolons: `Fixed Carrara slope matching; Repeated samples get their own groups`. This becomes the bullet list users see in the update prompt.
5. Press **Run workflow**. In about 15 minutes:
   - the app and installer are built on Windows and macOS, and each built app is launched to check it starts;
   - only if both platforms pass: `src/version.py` is bumped and committed, the commit is tagged `vX.Y.Z`, and the release is published with the app for each platform;
   - the **Installer** release gets fresh installers.

Every installed copy then sees the new version on its next launch. If any step fails, nothing is tagged or published — open the failed run in the Actions tab to see why, fix it, and run the workflow again.

Pull your `main` branch afterwards (`git pull`) to pick up the version-bump commit.

> The release files are named `MRSI.Data.Normalization.Tool.exe` and `MRSI.Data.Normalization.Tool.app.zip`. Installed apps look for exactly these names, so don't rename them.

---

## Building locally

Only needed to test a build by hand; releases are built by GitHub Actions.

```bash
pip install -r requirements-build.txt  # requirements.txt + PyInstaller

python packaging/build.py all          # app and installer -> dist/
python packaging/release.py smoke      # launch the built app, check it starts
python packaging/release.py package    # release-named files -> dist/release/
python packaging/build.py clean        # remove build/, dist/ and __pycache__
```

The spec picks the right shape for the host platform: a single self-contained `.exe` with a startup splash on Windows, and a `.app` bundle on macOS — the layout the in-app updater expects when it swaps a new version into place. Cross-compiling isn't possible; build each platform on that platform (which is what the workflow does).

`assets/images` is copied into the build by the spec and read back at runtime through `utils.resources`, which resolves paths against `sys._MEIPASS` when frozen.

---

## Project layout

```
.github/workflows/
├── checks.yml                     On every push: compile + start the app on Windows and macOS
└── release.yml                    Run by hand: build, test, version, tag and publish a release
assets/images/                     Logos and platform icons (single source of truth)
packaging/
├── build.py                       Build/clean entry point
├── release.py                     Release steps: next version, stamp, smoke test, package, notes
├── app.spec                       PyInstaller spec for the tool
├── installer.spec                 PyInstaller spec for the installer
└── entitlements.plist             macOS codesigning entitlements
tools/
├── generate_password_hash.py      Regenerate the Advanced Settings credentials
└── make_macos_icon.py             Rebuild logo.icns from the logo
requirements.txt                   Runtime dependencies
requirements-build.txt             + PyInstaller, for building
src/
├── main.py                        Entry point: splash, then staged startup
├── version.py                     The version number (stamped by the release workflow)
├── gui/
│   ├── splash.py                  Startup splash and update check
│   ├── main_window.py             Main window, tabs, theming, history, updates
│   ├── block_layout_bar.py        Carbonate / Yield / CO₂ block picker
│   └── tabs/
│       ├── carbonate_tab.py       Carbonate run configuration
│       ├── water_tab.py           Water run configuration
│       ├── combine_tab.py         Batch file list and combine options
│       ├── advanced_settings_tab.py   Settings, reset, and settings history
│       └── combine_processors/    Combine workers (one per mode)
├── processors/
│   ├── carbonate/step1…step7      Carbonate pipeline
│   └── water/step1…step7          Water pipeline
└── utils/
    ├── resources.py               Asset lookup, in a checkout and when frozen
    ├── settings.py                Runtime configuration, defaults, history, import/export
    ├── colors.py                  Reference-material colors (UI and Excel)
    ├── excel_engine.py            Shared background Excel instance
    ├── common_utils.py            Atomic saves, embedded settings comments
    ├── calculators/               Isotope fractionation maths
    ├── updater.py                 Release check, verified download, self-update
    └── installer/                 Standalone installer app
```

---

## Troubleshooting

**"Sheet 'X_DNT' not found. Run Step N first."**
Steps depend on the sheet the previous one creates. Either tick the earlier step or run the full chain.

**"Required column(s) missing from sheet …"**
The raw export doesn't have a column the pipeline groups on (`Identifier 1` and `Time Code` for carbonate, `Line` for water). The message lists what it did find.

**"… is a sheet this tool generates, not raw data."**
The sheet name box is pointing at a `_DNT` sheet. Pick the original instrument sheet — the tool auto-detects the right-most sheet when you drop a file.

**Run Selected Steps is greyed out**
Hover over it: it says what's missing — usually a file, or at least one ticked step.

**Excel windows flashing during processing**
Expected. Each step opens the workbook to recalculate formulas, then closes it. On macOS you may see Excel bounce in the Dock.

**`The RPC server is unavailable` (Windows)**
The background Excel instance was killed mid-run. The tool restarts Excel and retries automatically; if you still see it, check whether antivirus or a cleanup tool is terminating `EXCEL.EXE`.

**An update failed**
Your existing version keeps working and reopens. On Windows a note explains where the downloaded file was kept; the usual cause is the app living in a folder that needs administrator rights. Reinstalling with the installer into a folder you own (e.g. Desktop or Documents) fixes it.

---

<div align="center">
  <sub>Built by Ibrahim Parvez for the McMaster Research Group for Stable Isotopologues</sub>
</div>
