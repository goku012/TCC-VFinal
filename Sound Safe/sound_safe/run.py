# run.py

import customtkinter as ctk
from sound_monitor.app import SoundMonitorApp

if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    app = SoundMonitorApp()
    app.mainloop()
