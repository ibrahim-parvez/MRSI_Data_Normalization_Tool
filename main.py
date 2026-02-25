import sys
import time
from PyQt6.QtWidgets import QApplication
from splash import StartupSplashScreen

def main():
    app = QApplication(sys.argv)
    
    # 1. Show Splash immediately
    splash = StartupSplashScreen()
    splash.show()
    
    # ---------------- SETTINGS ----------------
    # The splash screen will stay open for at LEAST this many seconds.
    MIN_DURATION = 1.5
    # ------------------------------------------

    # Helper function to smooth out the jumps
    def smooth_progress(current_val, target_val, time_allocated, task_start_time):
        """
        Animates from current_val to target_val using whatever time is left 
        in the 'time_allocated' budget.
        """
        # How much time actually passed during the import/task?
        time_spent = time.time() - task_start_time
        
        # How much time do we have left to kill to meet the smooth quota?
        time_left = time_allocated - time_spent
        
        # Calculate the distance to travel on the bar
        steps_to_move = target_val - current_val
        
        if time_left > 0 and steps_to_move > 0:
            # We have extra time! Animate smoothly.
            delay_per_step = time_left / steps_to_move
            
            # Don't let the animation be too slow (cap at 0.05s per % so it doesn't lag)
            for i in range(current_val + 1, target_val + 1):
                splash.update_progress(i)
                app.processEvents()
                # If delay is tiny, sleep less or skip, but this is safe for >3s durations
                time.sleep(delay_per_step)
        else:
            # We are running late (computer is slow), just jump instantly
            splash.update_progress(target_val)
            app.processEvents()

    # Calculate roughly how much time each major step should take visually
    # We have 3 main blocks, so divide duration by 3
    chunk_time = MIN_DURATION / 3.0 

    # --- START LOADING ---
    splash.update_progress(0)
    
    # ==========================================
    # STEP 1: LOAD DATA ENGINES (Pandas)
    # ==========================================
    step_start = time.time()
    splash.loading_text.setText("Loading Data Engines (Pandas)...")
    app.processEvents()
    
    import pandas  # Blocking import
    
    # Smoothly move from 0 -> 33% using remaining time
    smooth_progress(0, 33, chunk_time, step_start)

    # ==========================================
    # STEP 2: LOAD GUI MODULES
    # ==========================================
    step_start = time.time()
    splash.loading_text.setText("Loading Interface Modules...")
    app.processEvents()
    
    from gui import DataToolApp # Blocking import
    
    # Smoothly move from 33% -> 66%
    smooth_progress(33, 66, chunk_time, step_start)

    # ==========================================
    # STEP 3: CONSTRUCT WINDOW
    # ==========================================
    step_start = time.time()
    splash.loading_text.setText("Constructing User Interface...")
    app.processEvents()
    
    window = DataToolApp() # Init the main window
    
    # Smoothly move from 66% -> 95%
    smooth_progress(66, 95, chunk_time, step_start)

    # ==========================================
    # FINISH
    # ==========================================
    splash.update_progress(100)
    splash.loading_text.setText("Ready!")
    app.processEvents()
    
    # Tiny pause at 100% so the user registers "Ready!" before it vanishes
    time.sleep(0.3)

    splash.close()
    window.show()

    sys.exit(app.exec())

if __name__ == "__main__":
    main()