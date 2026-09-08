"""Entry point for the MRSI Data Normalization Tool.

Shows the splash screen first, then imports the heavy modules (pandas, the
main window) while the progress bar advances, so the app appears instantly
instead of after a multi-second cold import.

The imports run on a worker thread rather than the GUI thread. macOS shows the
spinning-beachball cursor as soon as an app stops servicing its event loop for
a couple of seconds, and a cold ``import pandas`` out of a frozen bundle takes
longer than that, so doing it inline made a healthy startup look like a hang.
"""

import ctypes
import sys
import time

from PyQt6.QtCore import QObject, QThread, QTimer, pyqtSignal
from PyQt6.QtWidgets import QApplication, QMessageBox

from gui.splash import StartupSplashScreen
from utils.resources import app_icon

# Kept at module level so the window is not garbage collected once main() returns.
window = None

# Likewise for the loader thread: a QThread that goes out of scope while still
# running takes the process down with it.
loader = None

# The splash screen stays up for at least this many seconds.
MIN_DURATION = 1.5

# How often the progress bar repaints while the worker thread imports.
PROGRESS_TICK_MS = 30

# Windows groups taskbar buttons by this id; without it the app inherits
# Python's own icon instead of ours.
APP_USER_MODEL_ID = "mrsi.dnt.1.0"


def _configure_windows_taskbar() -> None:
    if sys.platform == "win32":
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)


def _close_bootloader_splash() -> None:
    """Dismiss the PyInstaller splash image used by the Windows one-file build.

    Only that build defines a splash, but ``pyi_splash`` is importable from
    any frozen build — elsewhere it complains on import and then raises on
    use, because the bootloader never initialised it. Hence the platform
    check, and the catch-all behind it: a decoration is never worth failing
    startup over.
    """
    if sys.platform != "win32":
        return

    try:
        import pyi_splash  # type: ignore

        pyi_splash.update_text("UI Loaded ...")
        pyi_splash.close()
    except Exception:
        pass


class ModuleLoader(QThread):
    """Imports the slow modules without blocking the GUI thread.

    Only imports happen here. The widgets themselves are still constructed on
    the main thread, because Qt requires it.
    """

    stage_started = pyqtSignal(str, int)
    loaded = pyqtSignal()
    failed = pyqtSignal(str)

    def run(self) -> None:
        try:
            self.stage_started.emit("Loading Data Engines (Pandas)...", 45)
            import pandas  # noqa: F401

            self.stage_started.emit("Loading Interface Modules...", 80)
            from gui.main_window import DataToolApp  # noqa: F401

            self.loaded.emit()
        except Exception as exc:  # noqa: BLE001 - reported to the user below
            self.failed.emit(f"{type(exc).__name__}: {exc}")


class SplashProgress(QObject):
    """Eases the splash progress bar toward a target on the GUI thread.

    The worker thread announces where the bar should be heading; this timer
    does the moving. Because it runs off the event loop, the splash keeps
    painting (and macOS keeps seeing a responsive app) throughout the imports.
    """

    def __init__(self, splash: StartupSplashScreen) -> None:
        super().__init__()
        self._splash = splash
        self._value = 0.0
        self._target = 0.0
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)
        self._timer.start(PROGRESS_TICK_MS)

    def set_target(self, target: int) -> None:
        self._target = max(self._target, float(target))

    def finish(self) -> None:
        self._timer.stop()
        self._value = 100.0
        self._splash.update_progress(100)

    def _tick(self) -> None:
        if self._value >= self._target:
            return

        # Asymptotic ease with a floor, so the bar always visibly creeps
        # forward even when a stage runs long, and never overshoots.
        step = max(0.4, (self._target - self._value) * 0.06)
        self._value = min(self._target, self._value + step)
        self._splash.update_progress(int(self._value))


def _load_application(app: QApplication, splash: StartupSplashScreen) -> None:
    """Kick off the background imports and wire up what happens when they land.

    Runs only once the splash screen's update check has finished.
    """
    global loader

    started_at = time.monotonic()
    progress = SplashProgress(splash)

    def on_stage(message: str, target: int) -> None:
        splash.loading_text.setText(message)
        progress.set_target(target)

    def on_failed(message: str) -> None:
        progress.finish()
        splash.close()
        QMessageBox.critical(
            None,
            "Startup Failed",
            f"The application could not finish loading.\n\n{message}",
        )
        app.exit(1)

    def on_loaded() -> None:
        global window

        splash.loading_text.setText("Constructing User Interface...")
        progress.set_target(95)
        # Let the bar and the new message paint before the main thread goes
        # busy building widgets.
        app.processEvents()

        from gui.main_window import DataToolApp

        window = DataToolApp()

        progress.finish()
        splash.loading_text.setText("Ready!")

        # Honour the minimum splash duration without blocking the event loop.
        remaining_ms = int(max(0.0, MIN_DURATION - (time.monotonic() - started_at)) * 1000)
        QTimer.singleShot(max(remaining_ms, 200), lambda: _reveal(splash))

    loader = ModuleLoader()
    loader.stage_started.connect(on_stage)
    loader.loaded.connect(on_loaded)
    loader.failed.connect(on_failed)
    loader.start()


def _reveal(splash: StartupSplashScreen) -> None:
    splash.close()
    if window is not None:
        window.show()
        window.raise_()
        window.activateWindow()


def main() -> None:
    _configure_windows_taskbar()

    app = QApplication(sys.argv)
    app.setWindowIcon(app_icon())

    splash = StartupSplashScreen()
    splash.show()
    _close_bootloader_splash()

    splash.startup_ready.connect(lambda: _load_application(app, splash))

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
