# gauge.py

import tkinter as tk
from .colors import DISCORD_SURFACE, DISCORD_SUCCESS, DISCORD_WARN, DISCORD_ERROR

class Gauge(tk.Canvas):
    def __init__(self, master, size=250, min_db=40, max_db=95, **kwargs):
        super().__init__(master, width=size, height=size, bg=DISCORD_SURFACE, highlightthickness=0, **kwargs)
        self.size = size
        self.center = size // 2
        self.radius = size * 0.42
        self.value = 0.0
        self.dose = 0.0
        self.min_db = float(min_db)
        self.max_db = float(max_db)
        self._ramp_target_pct = None
        self._ramp_rate_per_sec = 1.0  # máx 1% por segundo (ajuste a gosto)
        self.ref_db = 85.0

    def set_value(self, value, dose):
        self.value = float(value)
        self.dose = max(0.0, min(1.0, float(dose)))
        self._draw()

    def set_bounds(self, min_db, max_db):
        self.min_db = float(min_db)
        self.max_db = float(max_db)
        self._draw()

    def set_profile_ref(self, ref_db):
        self.ref_db = float(ref_db)

    def _draw(self):
        self.delete("all")
        start = -210
        extent = 240

        # fundo
        self.create_arc(self.center - self.radius, self.center - self.radius,
                        self.center + self.radius, self.center + self.radius,
                        start=start, extent=extent, style="arc", width=20, outline="#444")
        
        safe_cut = self.ref_db - 15.0
        warn_cut = self.ref_db
        if self.value < safe_cut:
            color = DISCORD_SUCCESS
        elif self.value < warn_cut:
            color = DISCORD_WARN
        else:
            color = DISCORD_ERROR

        ratio = (self.value - self.min_db) / max(1e-9, (self.max_db - self.min_db))
        ratio = max(0.0, min(1.0, ratio))
        fill_extent = extent * ratio

        self.create_arc(self.center - self.radius, self.center - self.radius,
                        self.center + self.radius, self.center + self.radius,
                        start=start, extent=fill_extent, style="arc", width=20, outline=color)

        self.create_text(self.center, self.center - 15, text=f"{self.value:.1f} dB",
                         fill="white", font=("Segoe UI", 24, "bold"))
        self.create_text(self.center, self.center + 20, text=f"Dose: {self.dose*100:.0f}%",
                         fill="white", font=("Segoe UI", 14, "bold"))
