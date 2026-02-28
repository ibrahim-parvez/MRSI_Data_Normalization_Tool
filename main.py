import sys
import time
from PyQt6.QtWidgets import QApplication
from splash import StartupSplashScreen

# Define window at the module level so it doesn't get garbage collected
window = None

def main():
    global window
    app = QApplication(sys.argv)
    
    # 1. Show Splash immediately
    splash = StartupSplashScreen()
    splash.show()
    
    # ---------------- SETTINGS ----------------
    # The splash screen will stay open for at LEAST this many seconds.
    MIN_DURATION = 1.5
    # ------------------------------------------

    def start_heavy_loading():
        """This runs ONLY after the update check finishes."""
        global window

        # Helper function to smooth out the jumps
        def smooth_progress(current_val, target_val, time_allocated, task_start_time):
            time_spent = time.time() - task_start_time
            time_left = time_allocated - time_spent
            steps_to_move = target_val - current_val
            
            if time_left > 0 and steps_to_move > 0:
                delay_per_step = time_left / steps_to_move
                for i in range(current_val + 1, target_val + 1):
                    splash.update_progress(i)
                    app.processEvents()
                    time.sleep(delay_per_step)
            else:
                splash.update_progress(target_val)
                app.processEvents()

        chunk_time = MIN_DURATION / 3.0 
        splash.update_progress(0)
        
        # ==========================================
        # STEP 1: LOAD DATA ENGINES (Pandas)
        # ==========================================
        step_start = time.time()
        splash.loading_text.setText("Loading Data Engines (Pandas)...")
        app.processEvents()
        
        import pandas  # Blocking import
        smooth_progress(0, 33, chunk_time, step_start)

        # ==========================================
        # STEP 2: LOAD GUI MODULES
        # ==========================================
        step_start = time.time()
        splash.loading_text.setText("Loading Interface Modules...")
        app.processEvents()
        
        from gui import DataToolApp # Blocking import
        smooth_progress(33, 66, chunk_time, step_start)

        # ==========================================
        # STEP 3: CONSTRUCT WINDOW
        # ==========================================
        step_start = time.time()
        splash.loading_text.setText("Constructing User Interface...")
        app.processEvents()
        
        window = DataToolApp() # Init the main window
        smooth_progress(66, 95, chunk_time, step_start)

        # ==========================================
        # FINISH
        # ==========================================
        splash.update_progress(100)
        splash.loading_text.setText("Ready!")
        app.processEvents()
        time.sleep(0.3)

        splash.close()
        window.show()

    # 2. Connect the signal from the splash screen to our loading function
    splash.startup_ready.connect(start_heavy_loading)

    sys.exit(app.exec())

if __name__ == "__main__":
    main()