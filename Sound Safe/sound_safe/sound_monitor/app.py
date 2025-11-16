# app.py

import tkinter as tk
import customtkinter as ctk
import time
import threading
import pythoncom
import math
import platform
import json
from pathlib import Path
from ctypes import POINTER, cast
from tkinter import messagebox, filedialog
from datetime import datetime
from queue import Queue, Empty

from .colors import (
    DISCORD_BG,
    DISCORD_SURFACE,
    DISCORD_SURFACE_ALT,
    DISCORD_ACCENT,
    DISCORD_SUCCESS,
    DISCORD_WARN,
    DISCORD_ERROR,
    DISCORD_TEXT,
)

from .excel_support import (
    _OPENPYXL_AVAILABLE,
    Workbook,
    get_column_letter,
    Font,
    Alignment,
    numbers,
)

from .audio_support import (
    _PYCAW_AVAILABLE,
    AudioUtilities,
    IAudioEndpointVolume,
    CLSCTX_ALL,
)

from .helpers import (
    map_percent_to_db,
    db_to_percent,
    allowed_time_seconds_for_level,
    dose_increment_per_second,
    risk_zone_from_dose,
    risk_zone_from_level,
    fmt_hms,
    round_pct_ui,
)

from .gauge import Gauge


# ---------- App ----------
class SoundMonitorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Monitor de Exposição Sonora - TCC")
        self.geometry("980x740"); self.minsize(860, 770)
        self.configure(fg_color=DISCORD_BG)

        # COM na thread da UI (PyCAW)
        self._ui_com_inited = False
        try:
            pythoncom.CoInitialize()
            self._ui_com_inited = True
        except Exception:
            pass

        # ===== Config DIÁRIA (8h) – perfis OMS/NIOSH =====
        # Estes valores são DIÁRIOS (8h). Troca 3 dB para ambos.
        self._defaults_cfg = {
            "min_db": 40.0,
            "max_db": 95.0,
            # padrão inicial: NIOSH 85/8h 3 dB
            "ref_db": 85.0,
            "base_time_sec": 8 * 3600.0,
            "exchange_rate_db": 3.0,

            "min_enforced_volume": 5.0,
            "default_volume": 30.0,
        }
        self.cfg = dict(self._defaults_cfg)

        # Preferências
        self.hard_lock_enabled = True
        self.lock_on_autoadjust = True

        # Estado de bloqueio
        self.locked = False
        self.lock_target_pct = None
        self.lock_reason = ""

        # Modos
        self.mode = "prefixado"
        self.dynamic_strategy = "reserva"  # 'reserva' | 'zona_segura'

        # Estado sessão / diária
        mynow = time.time()
        self.session_dose = 0.0   # dose relativa ao dia (base 8h de ref)
        self.prev_session_dose = 0.0
        self.time_at_current_level = 0.0

        self.daily_dose = 0.0
        now = datetime.now()
        self._day_key = now.strftime("%Y-%m-%d")

        self.alert_50_fired = False
        self.alert_100_fired = False
        self.daily_warn_fired = False
        self.daily_block_fired = False

        self._last_update = mynow
        self._stop_event = threading.Event()

        # Timer "neste volume"
        self._last_L_for_timer = None
        self.timer_epsilon_db = 1.0
        self._last_vol_key = None

        # Histórico / gráfico
        self.history = []
        self.session_start_ts = mynow
        self._last_hist_log = 0.0
        self._last_chart_draw = 0.0
        self.chart_window_sec = 120
        self.chart_points = []
        self._last_sys_sync = 0.0

        # Áudio backend
        self._audio_volume = None
        self._audio_warned = False
        self._init_audio_backend()

        # threads auxiliares de bloqueio
        self._lock_enforcer_thread = None
        self._lock_enforcer_stop = threading.Event()

        # UI flags
        self._slider_updating = False
        self.paused = False

        # Dinâmico (anti-oscilação)
        self.dynamic_reserve_min_sec = 600.0     # 10 min
        self.dynamic_reserve_max_sec = 1200.0    # 20 min
        self.dynamic_reserve_fraction = 0.10
        self.dynamic_step_small = 0.25
        self.dynamic_step_medium = 0.5
        self.dynamic_step_large = 1.0
        self.dynamic_hysteresis_sec = 90.0
        self.dynamic_adjust_interval = 0.6
        self.dynamic_limiting_active = False
        self.dynamic_decay_active = False
        self.last_dynamic_adjust_ts = 0.0

        # ----- Soft-lock dinâmico (teto móvel) -----
        self.dynamic_softlock_enabled = True     # trava aumentos enquanto o Dinâmico reduz
        self.dynamic_ceiling_pct = None          # teto atual (None = liberado)
        self.dynamic_release_delay = 20.0        # seg acima do 'upper' para liberar teto
        self._dynamic_upper_ok_since = None      # timestamp quando passou do upper

        # Quantização (combina com mixer do Windows)
        self._volume_quantum = 1.0  # % (1.0 se quiser mais fino)

        # Suavização EMA do “tempo restante”
        self._ema_remaining_sec = None
        self._ema_alpha = 0.25

        # UI
        self.left_frame = ctk.CTkFrame(self, width=200, corner_radius=10, fg_color=DISCORD_SURFACE)
        self.left_frame.pack(side="left", fill="y", padx=8, pady=8)
        self.right_frame = ctk.CTkFrame(self, corner_radius=10, fg_color=DISCORD_SURFACE)
        self.right_frame.pack(side="right", fill="both", expand=True, padx=8, pady=8)
        self._build_left_panel()
        self._build_right_panel()

        # Cache do slider
        self._vol_cache = float(self.vol_slider.get())

        # Fila de UI
        self._ui_queue = Queue()
        self.after(20, self._ui_pump)

        # Carrega settings
        self._load_settings()

        # Thread de monitoramento
        self._start_monitor_thread()

        # Sync inicial com SO
        if self._audio_volume is not None:
            try:
                sv = self._get_system_volume_percent()
                if abs(sv - float(self.vol_slider.get())) > 2.0:
                    self._safe_set_slider(sv)
                self.vol_label.configure(text=f"{round_pct_ui(self.vol_slider.get())}%")
                self._vol_cache = float(self.vol_slider.get())
            except Exception:
                pass

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------- Dispatcher de UI ----------
    def _ui_pump(self):
        try:
            while True:
                func = self._ui_queue.get_nowait()
                try:
                    func()
                except Exception as e:
                    print("Erro ao executar função de UI:", e)
        except Empty:
            pass
        self.after(20, self._ui_pump)

    def _on_ui(self, func):
        self._ui_queue.put(func)

    # ---------- Persistência ----------
    def _settings_path(self):
        base = Path.home() / ".tcc_sound_monitor"
        base.mkdir(parents=True, exist_ok=True)
        return base / "settings.json"

    def _load_settings(self):
        try:
            p = self._settings_path()
            if p.exists():
                with open(p, "r", encoding="utf-8") as fh:
                    data = json.load(fh)
                # cfg
                if isinstance(data.get("cfg"), dict):
                    for k, v in self._defaults_cfg.items():
                        self.cfg[k] = data["cfg"].get(k, v)
                # preferências
                self.hard_lock_enabled = bool(data.get("hard_lock_enabled", True))
                self.lock_on_autoadjust = bool(data.get("lock_on_autoadjust", True))
                # soft-lock
                self.dynamic_softlock_enabled = bool(data.get("dynamic_softlock_enabled", True))
                # estratégia dinâmico
                self.dynamic_strategy = data.get("dynamic_strategy", "reserva")
                if self.dynamic_strategy not in ("reserva", "zona_segura"):
                    self.dynamic_strategy = "reserva"
                # modo
                mode = data.get("mode")
                if mode in ("prefixado", "dinamico"):
                    self.set_mode(mode, silent=True)
                # volume
                vol = float(data.get("volume", self.cfg["default_volume"]))
                self._safe_set_slider(vol)
                self.vol_label.configure(text=f"{round_pct_ui(vol)}%")
                self._vol_cache = vol
            else:
                self.set_mode("prefixado", silent=True)
                self._safe_set_slider(self.cfg["default_volume"])
                self._vol_cache = float(self.cfg["default_volume"])
        except Exception as e:
            print("Falha ao carregar settings:", e)
        finally:
            self._refresh_profile_label()
            self.gauge.set_bounds(self.cfg["min_db"], self.cfg["max_db"])
            self.gauge.set_profile_ref(self.cfg["ref_db"])   # <--- adicione

    def _save_settings(self):
        try:
            data = {
                "mode": self.mode,
                "volume": float(self._vol_cache),
                "cfg": self.cfg,
                "hard_lock_enabled": self.hard_lock_enabled,
                "lock_on_autoadjust": self.lock_on_autoadjust,
                "dynamic_strategy": self.dynamic_strategy,
                "dynamic_softlock_enabled": self.dynamic_softlock_enabled,
            }
            with open(self._settings_path(), "w", encoding="utf-8") as fh:
                json.dump(data, fh, ensure_ascii=False, indent=2)
        except Exception as e:
            print("Falha ao salvar settings:", e)

    # ---------- Áudio (PyCAW) ----------
    def _init_audio_backend(self):
        if platform.system() != "Windows" or not _PYCAW_AVAILABLE:
            return
        try:
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            self._audio_volume = cast(interface, POINTER(IAudioEndpointVolume))
        except Exception:
            self._audio_volume = None

    def _get_system_volume_percent(self):
        if self._audio_volume is None:
            raise RuntimeError("Sem backend de áudio")
        scalar = self._audio_volume.GetMasterVolumeLevelScalar()
        return float(scalar) * 100.0

    def _set_system_volume_percent(self, pct):
        if self._audio_volume is None:
            raise RuntimeError("Sem backend de áudio")
        pct = max(0.0, min(100.0, float(pct)))
        self._audio_volume.SetMasterVolumeLevelScalar(pct / 100.0, None)

    # ---------- UI ----------
    def _build_left_panel(self):
        ctk.CTkLabel(self.left_frame, text="🎧", font=("Segoe UI", 64)).pack(pady=(10, 0))

        self.btn_prefixado = ctk.CTkButton(self.left_frame, text="Prefixado (ajuste imediato)",
                                           width=200, fg_color=DISCORD_ACCENT,
                                           command=lambda: self.set_mode("prefixado"))
        self.btn_prefixado.pack(pady=(12, 6))

        self.btn_dinamico = ctk.CTkButton(self.left_frame, text="Dinâmico (auto-limitando)",
                                          width=200, fg_color="#444",
                                          command=lambda: self.set_mode("dinamico"))
        self.btn_dinamico.pack(pady=6)

        self.mode_info = ctk.CTkLabel(self.left_frame, text="—", text_color="#bbb",
                                      wraplength=180, justify="left")
        self.mode_info.pack(pady=(8, 4))

        ctk.CTkLabel(self.left_frame, text="Dev: Breno Landim", font=("Segoe UI", 12),
                     text_color="#aaa").pack(side="bottom", pady=10)

    def _build_right_panel(self):
        ctk.CTkLabel(self.right_frame, text="Status de Exposição", font=("Segoe UI", 26, "bold")).pack(pady=8)

        status_frame = ctk.CTkFrame(self.right_frame, fg_color=DISCORD_SURFACE)
        status_frame.pack(fill="both", expand=True, padx=20, pady=10)

        gauge_frame = ctk.CTkFrame(status_frame, fg_color=DISCORD_SURFACE)
        gauge_frame.pack(side="left", fill="both", expand=True)
        self.gauge = Gauge(gauge_frame, size=220, min_db=self.cfg["min_db"], max_db=self.cfg["max_db"])
        self.gauge.pack(pady=10)

        info_frame = ctk.CTkFrame(status_frame, fg_color=DISCORD_SURFACE)
        info_frame.pack(side="right", fill="both", expand=True, padx=20, pady=10)

        self.zone_text_label = tk.Label(info_frame, text="Zona", font=("Segoe UI", 12, "bold"),
                                        bg=DISCORD_SURFACE, fg="white")
        self.zone_text_label.pack(pady=(0, 0))

        self.zone_canvas = tk.Canvas(info_frame, width=160, height=64, bg=DISCORD_SURFACE, highlightthickness=0)
        self.zone_canvas.pack(pady=(0, 0), expand=True)

        def draw_zone_badge(text, color):
            self.zone_canvas.delete("all")
            r = 16; x1, y1, x2, y2 = 0, 0, 160, 64
            self.zone_canvas.create_arc(x1, y1, x1 + 2*r, y1 + 2*r, start=90, extent=90, fill=color, outline=color)
            self.zone_canvas.create_arc(x2 - 2*r, y1, x2, y1 + 2*r, start=0, extent=90, fill=color, outline=color)
            self.zone_canvas.create_arc(x1, y2 - 2*r, x1 + 2*r, y2, start=180, extent=90, fill=color, outline=color)
            self.zone_canvas.create_arc(x2 - 2*r, y2 - 2*r, x2, y2, start=270, extent=90, fill=color, outline=color)
            self.zone_canvas.create_rectangle(x1 + r, y1, x2 - r, y2, fill=color, outline=color)
            self.zone_canvas.create_rectangle(x1, y1 + r, x2, y2 - r, fill=color, outline=color)
            self.zone_canvas.create_text(80, 32, text=text, font=("Segoe UI", 18, "bold"), fill="white")
        self.draw_zone_badge = draw_zone_badge
        self.draw_zone_badge("SEGURA", DISCORD_SUCCESS)

        self.time_label = tk.Label(self.right_frame, text="Tempo permitido: --:--:-- | Tempo neste volume: --:--:--",
                                   font=("Segoe UI", 14), bg=DISCORD_SURFACE, fg="white")
        self.time_label.pack(pady=(6, 0))

        self.remaining_label = tk.Label(self.right_frame, text="Tempo restante (neste volume) até 100%: --:--:--",
                                        font=("Segoe UI", 14), bg=DISCORD_SURFACE, fg="white")
        self.remaining_label.pack(pady=(2, 6))

        self.general_status = tk.Label(self.right_frame, text="Status: normal",
                                       font=("Segoe UI", 12, "bold"), bg=DISCORD_SURFACE, fg="#bbb")
        self.general_status.pack(pady=(0, 6))

        self.profile_label = tk.Label(self.right_frame, text=self._format_profile_text(),
                                      font=("Segoe UI", 11), bg=DISCORD_SURFACE, fg="#9aa0a6")
        self.profile_label.pack(pady=(0, 8))

        self.period_label = tk.Label(self.right_frame, text="Dose diária: 0%",
                                     font=("Segoe UI", 12), bg=DISCORD_SURFACE, fg="#bbb")
        self.period_label.pack(pady=(0, 10))

        vol_frame = ctk.CTkFrame(self.right_frame, fg_color=DISCORD_SURFACE)
        vol_frame.pack(pady=10)
        tk.Label(vol_frame, text="🔊", font=("Segoe UI Emoji", 15), bg=DISCORD_SURFACE, fg="white").pack(side="left", padx=5)

        self.vol_slider = ctk.CTkSlider(vol_frame, from_=0, to=100, number_of_steps=200,
                                        command=self.on_vol_slider_change, width=520, height=28,
                                        progress_color=DISCORD_ACCENT)
        self.vol_slider.set(50)
        self.vol_slider.pack(side="left", padx=10)

        self.vol_label = tk.Label(vol_frame, text=f"{round_pct_ui(50)}%", font=("Segoe UI", 10, "bold"),
                                  bg=DISCORD_SURFACE, fg="white", width=5)
        self.vol_label.pack(side="left", padx=10)

        chart_frame = ctk.CTkFrame(self.right_frame, fg_color=DISCORD_SURFACE)
        chart_frame.pack(fill="x", padx=20, pady=(8, 0))
        self.chart_canvas = tk.Canvas(chart_frame, width=760, height=120, bg=DISCORD_SURFACE_ALT, highlightthickness=0)
        self.chart_canvas.pack()

        btn_frame = ctk.CTkFrame(self.right_frame, fg_color=DISCORD_SURFACE)
        btn_frame.pack(side="bottom", pady=18)

        self.btn_reset = ctk.CTkButton(btn_frame, text="Resetar sessão",
                                       width=120, height=56, fg_color="#444",
                                       font=("Segoe UI", 14, "bold"),
                                       command=self.reset_session)
        self.btn_reset.pack(side="left", padx=10, pady=10)

        self.btn_excel = ctk.CTkButton(btn_frame, text="Salvar Relatório (Excel)",
                                       width=180, height=56, fg_color="#0078D7",
                                       font=("Segoe UI", 14, "bold"),
                                       command=self.save_report)
        self.btn_excel.pack(side="left", padx=10, pady=10)

        self.btn_cfg = ctk.CTkButton(btn_frame, text="Configurações",
                                     width=120, height=56, fg_color=DISCORD_ACCENT,
                                     font=("Segoe UI", 14, "bold"),
                                     command=self._open_settings_modal)
        self.btn_cfg.pack(side="left", padx=10, pady=10)

        self.pause_btn = ctk.CTkButton(btn_frame, text="Pausar",
                                       width=150, height=56, fg_color="#6b7280",
                                       font=("Segoe UI", 14, "bold"),
                                       command=self._toggle_pause)
        self.pause_btn.pack(side="left", padx=10, pady=10)

    def _quantize_pct(self, pct: float) -> float:
        q = float(self._volume_quantum) if getattr(self, "_volume_quantum", None) else 1.0
        return max(0.0, min(100.0, round(float(pct) / q) * q))

    # ---------- Helpers UI ----------
    def _format_profile_text(self):
        hours = self.cfg["base_time_sec"] / 3600.0  # 8h (fixo)
        er = int(self.cfg["exchange_rate_db"]) if float(self.cfg["exchange_rate_db"]).is_integer() else self.cfg["exchange_rate_db"]
        return f"Perfil diário: {self.cfg['ref_db']:.0f} dB / {hours:g}h ({er} dB)"

    def _refresh_profile_label(self):
        self.profile_label.config(text=self._format_profile_text())

    def _safe_set_slider(self, pct):
        pct = self._quantize_pct(pct)
        self._vol_cache = float(pct)
        try:
            self._slider_updating = True
            self.vol_slider.set(pct)
        finally:
            self._slider_updating = False

    # ---------- Bloqueio ----------
    def _lock_volume(self, target_pct: float, reason: str = "", honor_min: bool = True):
        self.locked = True
        target = float(target_pct)
        if honor_min:
            target = max(self.cfg["min_enforced_volume"], target)

        self.lock_target_pct = target
        self.lock_reason = reason
        self.vol_slider.configure(state="disabled")
        self.btn_dinamico.configure(state="disabled")
        self.btn_prefixado.configure(state="disabled")
        self.pause_btn.configure(state="disabled")
        self._safe_set_slider(self.lock_target_pct)
        self._apply_system_volume_from_slider(show_install_hint=True)
        self.general_status.config(text=f"Status: bloqueado ({reason})", fg=DISCORD_ERROR)
        self._start_lock_enforcer()

    def _unlock_volume(self):
        self.locked = False
        self.lock_target_pct = None
        self.lock_reason = ""
        self._stop_lock_enforcer()
        self.vol_slider.configure(state="normal")
        self.btn_dinamico.configure(state="normal")
        self.btn_prefixado.configure(state="normal")
        self.pause_btn.configure(state="normal")
        self.general_status.config(text="Status: normal", fg="#bbb")

    # ---------- Modo ----------
    def set_mode(self, mode, silent=False):
        if mode not in ("prefixado", "dinamico"):
            return
        self.mode = mode
        self.time_at_current_level = 0.0
        self._last_L_for_timer = None
        self._last_vol_key = None
        self.dynamic_limiting_active = False
        self.dynamic_decay_active = False
        self.last_dynamic_adjust_ts = 0.0

        # reset do soft-lock dinâmico
        self.dynamic_ceiling_pct = None
        self._dynamic_upper_ok_since = None

        def set_btn_colors(p="#444", d="#444"):
            self.btn_prefixado.configure(fg_color=p)
            self.btn_dinamico.configure(fg_color=d)

        if mode == "prefixado":
            set_btn_colors(p=DISCORD_ACCENT)
            self.mode_info.configure(text="Passou do limite? Ajuste imediato para o volume seguro (mantém ≥10 min de folga).")
            self._unlock_volume()  # <- adicione esta linha
        else:
            set_btn_colors(d=DISCORD_ACCENT)
            if self.dynamic_strategy == "reserva":
                self.mode_info.configure(text="Dinâmico (Reserva): mantém 10–20 min de folga e reduz suavemente quando precisa.")
            else:
                self.mode_info.configure(text="Dinâmico (Zona Segura): reduz gradualmente até entrar no verde do gauge.")

        if not silent:
            self.general_status.config(text="Status: normal", fg="#bbb")
        self._refresh_profile_label()

    # ---------- Slider ----------
    def on_vol_slider_change(self, value):
        if self.locked:
            self._safe_set_slider(self.lock_target_pct)
            self._apply_system_volume_from_slider(show_install_hint=False)
            return
        if getattr(self, "_slider_updating", False):
            return

        v = float(value)

        # Não permitir subir enquanto dinâmica está descendo
        if self.dynamic_decay_active and v > self._vol_cache + 0.01:
            self._safe_set_slider(self._vol_cache)
            return

        # Soft-lock: impede subir acima do teto enquanto o Dinâmico estiver atuando
        if self.dynamic_softlock_enabled and self.dynamic_ceiling_pct is not None and v > self.dynamic_ceiling_pct + 0.01:
            self._safe_set_slider(self.dynamic_ceiling_pct)
            self._apply_system_volume_from_slider(show_install_hint=False)
            return

        self._vol_cache = v
        self.vol_label.configure(text=f"{round_pct_ui(v)}%")
        self._apply_system_volume_from_slider(show_install_hint=True)

    # ---------- Ações ----------
    def _toggle_pause(self):
        self.paused = not self.paused
        if self.paused:
            self.general_status.config(text="Status: pausado", fg=DISCORD_WARN)
            self.pause_btn.configure(text="Retomar")
        else:
            self.general_status.config(text="Status: normal", fg="#bbb")
            self.pause_btn.configure(text="Pausar")

    def reset_session(self):
        self.session_dose = 0.0
        self.prev_session_dose = 0.0
        self.alert_50_fired = False
        self.alert_100_fired = False
        self.time_at_current_level = 0.0
        self.session_start_ts = time.time()
        self.history = []
        self.chart_points = []
        self._last_hist_log = 0.0
        self._last_chart_draw = 0.0
        self._last_update = time.time()
        self._last_L_for_timer = None
        self._last_vol_key = None
        self.dynamic_limiting_active = False
        self.dynamic_decay_active = False
        self.last_dynamic_adjust_ts = 0.0
        self.dynamic_ceiling_pct = None
        self._dynamic_upper_ok_since = None
        self.general_status.config(text="Status: normal", fg="#bbb")
        self._on_ui(lambda: self.remaining_label.config(text="Tempo restante (neste volume) até 100%: --:--:--"))
        self._unlock_volume()
        messagebox.showinfo("Sessão reiniciada", "Dose e histórico foram resetados.")

    # ---------- Exportar Excel ----------
    def save_report(self):
        if not _OPENPYXL_AVAILABLE:
            messagebox.showerror("Dependência ausente", "Para exportar Excel (.xlsx): pip install openpyxl")
            return
        if not self.history and not messagebox.askyesno("Sem dados", "Ainda não há histórico. Salvar mesmo assim?"):
            return
        filename = filedialog.asksaveasfilename(
            title="Salvar relatório Excel",
            defaultextension=".xlsx",
            filetypes=[("Pasta de trabalho do Excel", "*.xlsx")],
            initialfile=f"relatorio_som_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        if not filename:
            return
        try:
            wb = Workbook()
            ws = wb.active; ws.title = "Relatório"
            headers = ["timestamp_iso","t_sessao_s","modo","volume_%","nivel_dB","dose_0a1","zona","dose_diaria"]
            ws.append(headers)
            for c in range(1, len(headers)+1):
                cell = ws.cell(row=1, column=c); cell.font = Font(bold=True); cell.alignment = Alignment(horizontal="center")
            for row in self.history:
                ws.append([
                    row["ts_iso"], float(row["t_session"]), row["mode"],
                    int(round_pct_ui(float(row["vol_percent"]))),
                    float(row["L"]), float(row["dose"]), row["zone"], float(row["daily"])
                ])
            for r in range(2, ws.max_row + 1):
                ws.cell(r, 2).number_format = "0.0"
                ws.cell(r, 5).number_format = "0.00"
                ws.cell(r, 6).number_format = numbers.FORMAT_PERCENTAGE_00
                ws.cell(r, 8).number_format = numbers.FORMAT_PERCENTAGE_00
            widths = [20,14,12,12,12,12,16,16]
            for idx, w in enumerate(widths, start=1):
                ws.column_dimensions[get_column_letter(idx)].width = w
            ws.auto_filter.ref = f"A1:H{ws.max_row}"; ws.freeze_panes = "A2"

            ws2 = wb.create_sheet(title="Resumo")
            if not self.history:
                ws2["A1"] = "Sem dados na sessão."; ws2["A1"].font = Font(bold=True)
            else:
                summary = self._compute_summary_stats()
                ws2["A1"] = "Resumo da Sessão"; ws2["A1"].font = Font(bold=True)
                ws2["A2"] = f"Gerado em: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
                ws2["A3"] = f"Perfil diário: {self.cfg['ref_db']:.0f} dB / 8h (3 dB)"
                labels = [
                    ("Tempo total",          summary["total_time_days"], "[h]:mm:ss"),
                    ("Média de dB",          summary["avg_db"],          "0.00"),
                    ("Pico de dB",           summary["peak_db"],         "0.00"),
                    ("Pico de volume (%)",   summary["peak_vol"],        "0"),
                    ("Maior dose (sessão)",  summary["max_dose"],        numbers.FORMAT_PERCENTAGE_00),
                    ("Tempo até 50% dose",   summary["t_to_50_days"],    "[h]:mm:ss"),
                    ("Tempo até 100% dose",  summary["t_to_100_days"],   "[h]:mm:ss"),
                ]
                ws2["A5"] = "Métricas gerais"; ws2["A5"].font = Font(bold=True)
                row_i = 6
                for label, value, fmt in labels:
                    ws2.cell(row=row_i, column=1, value=label)
                    cval = ws2.cell(row=row_i, column=2, value=value); cval.number_format = fmt
                    row_i += 1
                ws2.column_dimensions["A"].width = 26
                ws2.column_dimensions["B"].width = 18
            wb.save(filename)
            messagebox.showinfo("Relatório salvo", f"Relatório Excel exportado em:\n{filename}")
        except Exception as e:
            messagebox.showerror("Erro ao salvar", f"Ocorreu um erro ao salvar o Excel:\n{e}")

    def _compute_summary_stats(self):
        hist = self.history; n = len(hist)
        total_time_s = 0.0; weighted_sum_L = 0.0
        peak_db = float("-inf"); peak_vol = float("-inf")
        max_dose = 0.0
        t_to_50 = None; t_to_100 = None
        if n >= 1:
            for i in range(n - 1):
                cur, nxt = hist[i], hist[i+1]
                dt = max(0.0, float(nxt["t_session"]) - float(cur["t_session"]))

                total_time_s += dt
                L = float(cur["L"]); weighted_sum_L += L * dt
                peak_db = max(peak_db, L)
                peak_vol = max(peak_vol, float(cur["vol_percent"]))
                max_dose = max(max_dose, float(cur["dose"]))

                if t_to_50 is None and float(cur["dose"]) >= 0.5: t_to_50 = float(cur["t_session"])
                if t_to_100 is None and float(cur["dose"]) >= 1.0: t_to_100 = float(cur["t_session"])

            last = hist[-1]
            peak_db = max(peak_db, float(last["L"]))
            peak_vol = max(peak_vol, float(last["vol_percent"]))
            max_dose = max(max_dose, float(last["dose"]))
            if t_to_50 is None and float(last["dose"]) >= 0.5: t_to_50 = float(last["t_session"])
            if t_to_100 is None and float(last["dose"]) >= 1.0: t_to_100 = float(last["t_session"])
        avg_db = (weighted_sum_L / total_time_s) if total_time_s > 0 else 0.0
        return {
            "points": n,
            "total_time_s": total_time_s,
            "total_time_days": total_time_s / 86400.0,
            "avg_db": avg_db,
            "peak_db": peak_db if peak_db != float("-inf") else 0.0,
            "peak_vol": peak_vol if peak_vol != float("-inf") else 0.0,
            "max_dose": max_dose,
            "t_to_50_days": (t_to_50 / 86400.0) if t_to_50 is not None else 0.0,
            "t_to_100_days": (t_to_100 / 86400.0) if t_to_100 is not None else 0.0,
        }

    # ---------- Teto (Prefixado) ----------
    def _calc_safe_zone_target_pct(self):
        # Zona segura = início do verde = ref_db - 15 dB
        Lmax = float(self.cfg["ref_db"]) - 15.0
        return db_to_percent(Lmax, self.cfg)

    # ---------- Gráfico ----------
    def _draw_history_chart(self):
        w = int(self.chart_canvas.winfo_width() or 760)
        h = int(self.chart_canvas.winfo_height() or 120)
        pad_l, pad_r, pad_t, pad_b = 40, 10, 10, 25
        self.chart_canvas.delete("all")
        self.chart_canvas.create_line(pad_l, h - pad_b, w - pad_r, h - pad_b, fill="#555")
        self.chart_canvas.create_line(pad_l, pad_t, pad_l, h - pad_b, fill="#555")
        if not self.chart_points:
            self.chart_canvas.create_text(w//2, h//2, text="Sem dados ainda", fill="#888", font=("Segoe UI", 10))
            return
        t_now = self.chart_points[-1][0]
        t_min = max(0.0, t_now - self.chart_window_sec)
        t_max = t_now
        span = max(1e-6, t_max - t_min)
        min_db = self.cfg["min_db"]; max_db = self.cfg["max_db"]
        def x_map(t): return pad_l + (w - pad_l - pad_r) * ((t - t_min) / span)
        def y_map_db(L):
            ratio = (L - min_db) / max(1e-9, (max_db - min_db))
            ratio = max(0.0, min(1.0, ratio))
            return (h - pad_b) - (h - pad_b - pad_t) * ratio
        def y_map_dose(d):
            ratio = max(0.0, min(1.0, float(d)))
            return (h - pad_b) - (h - pad_b - pad_t) * ratio
        last_x_db = last_y_db = last_x_ds = last_y_ds = None
        for (t_rel, L, dose) in self.chart_points:
            if t_rel < t_min: continue
            x = x_map(t_rel); y_db = y_map_db(L); y_ds = y_map_dose(dose)
            if last_x_db is not None:
                self.chart_canvas.create_line(last_x_db, last_y_db, x, y_db, fill="#8FD14F", width=2)
            last_x_db, last_y_db = x, y_db
            if last_x_ds is not None:
                self.chart_canvas.create_line(last_x_ds, last_y_ds, x, y_ds, fill="#4FC3F7", width=2)
            last_x_ds, last_y_ds = x, y_ds
        self.chart_canvas.create_text(w - 140, pad_t + 12, text="dB", fill="#8FD14F", font=("Segoe UI", 10, "bold"))
        self.chart_canvas.create_text(w - 90, pad_t + 12, text="Dose%", fill="#4FC3F7", font=("Segoe UI", 10, "bold"))
        for label, Lbl in [("min", min_db), ("ref", self.cfg["ref_db"]), ("max", max_db)]:
            y = y_map_db(Lbl)
            self.chart_canvas.create_line(pad_l - 5, y, w - pad_r, y, fill="#333")
            self.chart_canvas.create_text(pad_l - 28, y, text=f"{Lbl:.0f}", fill="#aaa", font=("Segoe UI", 9))
        for dt in range(0, int(self.chart_window_sec) + 1, 30):
            x = x_map(t_max - dt)
            self.chart_canvas.create_line(x, h - pad_b, x, pad_t, fill="#333")
            self.chart_canvas.create_text(x, h - pad_b + 12, text=f"-{dt}s", fill="#aaa", font=("Segoe UI", 9))

    # ---------- Monitor ----------
    def _start_monitor_thread(self):
        t = threading.Thread(target=self._monitor_loop, daemon=True)
        t.start()

    def _apply_system_volume_from_slider(self, show_install_hint=False):
        if self._audio_volume is not None:
            try:
                target = self._quantize_pct(self._vol_cache)
                self._set_system_volume_percent(target)
                return
            except Exception:
                pass
        if (platform.system() == "Windows") and show_install_hint and not self._audio_warned:
            self._audio_warned = True
            messagebox.showinfo(
                "Controlar volume do Windows",
                "Para o slider controlar (e travar) o volume do PC, instale: pip install pycaw comtypes"
            )

    def _start_lock_enforcer(self):
        if self._lock_enforcer_thread and self._lock_enforcer_thread.is_alive():
            return
        self._lock_enforcer_stop.clear()
        def _runner():
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass
            try:
                while self.locked and not self._lock_enforcer_stop.is_set():
                    try:
                        if self._audio_volume is not None and self.lock_target_pct is not None:
                            current = self._get_system_volume_percent()
                            target = float(self.lock_target_pct)
                            if abs(current - target) > 0.5:
                                self._set_system_volume_percent(target)
                    except Exception:
                        pass
                    time.sleep(0.03)
            finally:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
        self._lock_enforcer_thread = threading.Thread(target=_runner, daemon=True)
        self._lock_enforcer_thread.start()

    def _stop_lock_enforcer(self):
        self._lock_enforcer_stop.set()

    def _roll_day_if_needed(self, now_dt):
        day_key = now_dt.strftime("%Y-%m-%d")
        if day_key != self._day_key:
            self._day_key = day_key
            self.daily_dose = 0.0
            self.daily_warn_fired = False
            self.daily_block_fired = False
            self.alert_50_fired = False
            self.alert_100_fired = False
            self.session_dose = 0.0
            self._on_ui(lambda: messagebox.showinfo("Novo dia", "Dose diária reiniciada."))

    def _monitor_loop(self):
        pythoncom.CoInitialize()
        try:
            while not self._stop_event.is_set():
                try:
                    if self.locked:
                        if abs(float(self._vol_cache) - float(self.lock_target_pct or 0)) > 0.1:
                            self._on_ui(lambda: self._safe_set_slider(self.lock_target_pct))
                        self._apply_system_volume_from_slider(show_install_hint=False)

                    vol_percent = float(self._vol_cache)
                    now = time.time()
                    dt = now - self._last_update
                    dt = max(0.0, min(dt, 1.0))
                    self._last_update = now

                    now_dt = datetime.now()
                    self._roll_day_if_needed(now_dt)

                    # dB corrente
                    L_eff = map_percent_to_db(vol_percent, self.cfg)

                    if self.paused:
                        self._on_ui(lambda: self.gauge.set_value(L_eff, self.session_dose))
                        self._on_ui(lambda: self.general_status.config(text="Status: pausado", fg=DISCORD_WARN))
                        self._apply_system_volume_from_slider(show_install_hint=False)
                        time.sleep(0.1 if self.locked else 0.2)
                        continue

                    # "tempo neste volume"
                    vol_key = int(round_pct_ui(self._vol_cache))
                    if self._last_L_for_timer is None:
                        self._last_L_for_timer = L_eff
                        self._last_vol_key = vol_key
                    else:
                        changed_db = abs(L_eff - self._last_L_for_timer) >= self.timer_epsilon_db
                        changed_pct = (self._last_vol_key is None) or (self._last_vol_key != vol_key)
                        if changed_db or changed_pct:
                            self._last_L_for_timer = L_eff
                            self._last_vol_key = vol_key
                            self.time_at_current_level = 0.0

                    # Dose diária (base 8h)
                    self.prev_session_dose = self.session_dose
                    inc = dose_increment_per_second(L_eff, self.cfg) * dt
                    self.session_dose = min(1.0, self.session_dose + inc)
                    self.daily_dose = min(10.0, self.daily_dose + inc)

                    # Cronômetro do nível atual
                    if self.session_dose < 1.0:
                        self.time_at_current_level += dt
                    else:
                        if not self.locked and self.hard_lock_enabled:
                            self._on_ui(lambda: self._lock_volume(self.cfg["min_enforced_volume"], reason="limite diário"))
                        self.time_at_current_level = 0.0

                    # Tempos (permitido + restante até 100%)
                    allowed_sec = allowed_time_seconds_for_level(L_eff, self.cfg)
                    time_str = fmt_hms(allowed_sec)
                    time_cur_str = fmt_hms(self.time_at_current_level)
                    remaining_sec = (1.0 - self.session_dose) * allowed_sec if self.session_dose < 1.0 else 0.0

                    # EMA para suavizar “tempo restante”
                    if self._ema_remaining_sec is None:
                        self._ema_remaining_sec = remaining_sec
                    else:
                        a = self._ema_alpha
                        self._ema_remaining_sec = a * remaining_sec + (1 - a) * self._ema_remaining_sec
                    ema_remaining = self._ema_remaining_sec
                    remaining_str = fmt_hms(ema_remaining)

                    # Zonas
                    zone = risk_zone_from_dose(self.session_dose)
                    zone_color = DISCORD_SUCCESS if zone == "SEGURA" else DISCORD_WARN if zone == "ATENÇÃO" else DISCORD_ERROR
                    level_zone = risk_zone_from_level(L_eff, self.cfg)

                    # ----- Regras dos modos -----
                    if self.mode == "prefixado" and not self.locked:
                        # Quando o tempo permitido no nível atual zera, corta p/ zona segura
                        if self.session_dose < 1.0:
                            if ema_remaining <= 0.0:
                                target = self._calc_safe_zone_target_pct()
                                self._on_ui(lambda: self._safe_set_slider(min(self._vol_cache, target)))
                                self._apply_system_volume_from_slider(show_install_hint=False)
                                self._on_ui(lambda: self.general_status.config(
                                    text="Status: corte p/ zona segura (perfil)", fg=DISCORD_WARN))
                                if self.lock_on_autoadjust:
                                    # trava no seguro pós-corte
                                    self._on_ui(lambda: self._lock_volume(target, reason="corte automático (perfil)", honor_min=False))
                        else:
                            self._on_ui(lambda: self.general_status.config(text="Status: normal", fg="#bbb"))

                    elif self.mode == "dinamico" and self.session_dose < 1.0 and not self.locked:
                        if self.dynamic_strategy == "reserva":
                            # ===== Estratégia RESERVA =====
                            reserve_target = max(
                                self.dynamic_reserve_min_sec,
                                min(self.dynamic_reserve_max_sec, self.dynamic_reserve_fraction * allowed_sec)
                            )
                            lower = reserve_target - self.dynamic_hysteresis_sec
                            upper = reserve_target + self.dynamic_hysteresis_sec

                            if not self.dynamic_limiting_active and ema_remaining < lower:
                                self.dynamic_limiting_active = True
                                # ao entrar no limitando, captura teto inicial
                                if self.dynamic_softlock_enabled:
                                    cur = self._quantize_pct(self._vol_cache)
                                    self.dynamic_ceiling_pct = cur if self.dynamic_ceiling_pct is None else min(self.dynamic_ceiling_pct, cur)
                                    self._dynamic_upper_ok_since = None

                            elif self.dynamic_limiting_active and ema_remaining > upper:
                                self.dynamic_limiting_active = False

                            if self.dynamic_limiting_active:
                                if (now - self.last_dynamic_adjust_ts) >= self.dynamic_adjust_interval:
                                    deficit = reserve_target - ema_remaining
                                    if deficit < 60: step = self.dynamic_step_small
                                    elif deficit < 300: step = self.dynamic_step_medium
                                    else: step = self.dynamic_step_large

                                    def _decay():
                                        current = self._vol_cache
                                        target = self._quantize_pct(current - step)
                                        if target < current - 0.099:
                                            self.dynamic_decay_active = True
                                            self._safe_set_slider(target)
                                            # atualiza teto (monótono)
                                            if self.dynamic_softlock_enabled:
                                                self.dynamic_ceiling_pct = target if self.dynamic_ceiling_pct is None else min(self.dynamic_ceiling_pct, target)
                                    self._on_ui(_decay)
                                    self.last_dynamic_adjust_ts = now

                                self._on_ui(lambda: self.general_status.config(text="Status: auto-limitando", fg=DISCORD_WARN))
                            else:
                                self.dynamic_decay_active = False
                                # Liberação do teto só após estabilidade acima de 'upper'
                                if self.dynamic_softlock_enabled:
                                    if ema_remaining > upper:
                                        if self._dynamic_upper_ok_since is None:
                                            self._dynamic_upper_ok_since = now
                                        elif (now - self._dynamic_upper_ok_since) >= self.dynamic_release_delay:
                                            self.dynamic_ceiling_pct = None
                                    else:
                                        self._dynamic_upper_ok_since = None

                                self._on_ui(lambda: self.general_status.config(text="Status: normal", fg="#bbb"))

                        else:
                            # ===== Estratégia ZONA SEGURA =====
                            if level_zone != "SEGURA":
                                if not self.dynamic_limiting_active:
                                    self.dynamic_limiting_active = True
                                    # pegamos um teto inicial
                                    if self.dynamic_softlock_enabled:
                                        cur = self._quantize_pct(self._vol_cache)
                                        self.dynamic_ceiling_pct = cur if self.dynamic_ceiling_pct is None else min(self.dynamic_ceiling_pct, cur)
                                        self._dynamic_upper_ok_since = None

                                if (now - self.last_dynamic_adjust_ts) >= self.dynamic_adjust_interval:
                                    if L_eff < 90: step = self.dynamic_step_small
                                    else: step = self.dynamic_step_medium if L_eff < 95 else self.dynamic_step_large
                                    def _decay2():
                                        new_v = max(self.cfg["min_enforced_volume"], self._vol_cache - step)
                                        if abs(new_v - self._vol_cache) >= 0.1:
                                            self._safe_set_slider(new_v)
                                            if self.dynamic_softlock_enabled:
                                                self.dynamic_ceiling_pct = new_v if self.dynamic_ceiling_pct is None else min(self.dynamic_ceiling_pct, new_v)
                                    self._on_ui(_decay2)
                                    self.last_dynamic_adjust_ts = now
                                self._on_ui(lambda: self.general_status.config(
                                    text="Status: auto-limitando (até zona segura)", fg=DISCORD_WARN))
                            else:
                                # estamos na zona segura
                                self.dynamic_limiting_active = False
                                self.dynamic_decay_active = False
                                # Libera teto após estabilidade em verde por X s
                                if self.dynamic_softlock_enabled:
                                    if level_zone == "SEGURA":
                                        if self._dynamic_upper_ok_since is None:
                                            self._dynamic_upper_ok_since = now
                                        elif (now - self._dynamic_upper_ok_since) >= self.dynamic_release_delay:
                                            self.dynamic_ceiling_pct = None
                                    else:
                                        self._dynamic_upper_ok_since = None
                                self._on_ui(lambda: self.general_status.config(text="Status: normal", fg="#bbb"))

                    # ----- Alertas DIÁRIOS -----
                    daily_pct = self.daily_dose * 100.0
                    if daily_pct >= 80.0 and not self.daily_warn_fired and daily_pct < 100.0:
                        self.daily_warn_fired = True
                        self._on_ui(lambda: messagebox.showwarning("Atenção diária", "Dose diária ≥ 80%."))
                    if 80.0 <= daily_pct < 100.0:
                        self._on_ui(lambda: self.period_label.config(fg=DISCORD_WARN))
                    elif daily_pct >= 100.0:
                        self._on_ui(lambda: self.period_label.config(fg=DISCORD_ERROR))
                    else:
                        self._on_ui(lambda: self.period_label.config(fg="#bbb"))

                    if daily_pct >= 100.0 and not self.daily_block_fired:
                        self.daily_block_fired = True
                        self._on_ui(lambda: messagebox.showerror("Bloqueio diário", "Dose diária atingiu 100%. Volume mínimo imposto."))
                        if self.hard_lock_enabled:
                            self._on_ui(lambda: self._lock_volume(self.cfg["min_enforced_volume"], reason="limite diário"))

                    # Alertas de dose (sessão = diária)
                    if not self.alert_50_fired and self.session_dose >= 0.5:
                        self.alert_50_fired = True
                        self._on_ui(lambda: messagebox.showwarning("Atenção", "Você atingiu 50% da dose diária."))
                    if not self.alert_100_fired and self.session_dose >= 1.0:
                        self.alert_100_fired = True
                        self._on_ui(lambda: messagebox.showerror("Risco crítico", "Limite de dose diária ultrapassado!"))
                        if self.hard_lock_enabled:
                            self._on_ui(lambda: self._lock_volume(self.cfg["min_enforced_volume"], reason="limite diário"))

                    # ----- Atualiza UI -----
                    self._on_ui(lambda: self.gauge.set_value(L_eff, self.session_dose))
                    self._on_ui(lambda: self.draw_zone_badge(zone, zone_color))
                    self._on_ui(lambda: self.time_label.config(
                        text=f"Tempo permitido: {time_str} | Tempo neste volume: {time_cur_str}"
                    ))
                    self._on_ui(lambda: self.remaining_label.config(
                        text=f"Tempo restante (neste volume) até 100%: {remaining_str}"
                    ))
                    self._on_ui(lambda: self.vol_slider.configure(progress_color=zone_color))
                    self._on_ui(lambda: self.vol_label.configure(text=f"{round_pct_ui(self._vol_cache)}%"))
                    self._on_ui(lambda: self.period_label.config(text=f"Dose diária: {daily_pct:.0f}%"))

                    # Sync com sistema periodicamente
                    if self._audio_volume is not None and (now - self._last_sys_sync) >= 0.5:
                        self._last_sys_sync = now
                        try:
                            sys_pct = self._quantize_pct(self._get_system_volume_percent())

                            # Se há teto, rebaixa o volume do Windows caso tenha subido acima dele
                            if self.dynamic_softlock_enabled and self.dynamic_ceiling_pct is not None:
                                if sys_pct > self.dynamic_ceiling_pct + 0.5:
                                    self._on_ui(lambda: self._safe_set_slider(self.dynamic_ceiling_pct))
                                    self._apply_system_volume_from_slider(show_install_hint=False)
                                    sys_pct = self.dynamic_ceiling_pct

                            if self.locked:
                                if abs(sys_pct - float(self.lock_target_pct or 0)) > 0.5:
                                    self._on_ui(lambda: self._safe_set_slider(self.lock_target_pct))
                                    self._apply_system_volume_from_slider(show_install_hint=False)
                            else:
                                # Se estivermos em queda dinâmica, ignora subidas externas
                                if self.dynamic_decay_active and sys_pct > float(self._vol_cache) + 0.01:
                                    sys_pct = self._vol_cache
                                if abs(sys_pct - float(self._vol_cache)) > 1.0:
                                    self._on_ui(lambda v=sys_pct: self._safe_set_slider(v))
                        except Exception:
                            pass

                    # Histórico (~1s)
                    if (now - self._last_hist_log) >= 1.0:
                        self._last_hist_log = now
                        t_rel = now - self.session_start_ts
                        self.history.append({
                            "ts_iso": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(now)),
                            "t_session": t_rel,
                            "mode": self.mode,
                            "vol_percent": float(self._vol_cache),
                            "L": L_eff,
                            "dose": self.session_dose,
                            "zone": zone,
                            "daily": self.daily_dose,
                        })
                        self.chart_points.append((t_rel, L_eff, self.session_dose))
                        cutoff = t_rel - self.chart_window_sec - 2
                        self.chart_points = [p for p in self.chart_points if p[0] >= cutoff]

                    # Redesenha gráfico (~0.8s)
                    if (now - self._last_chart_draw) >= 0.8:
                        self._last_chart_draw = now
                        self._on_ui(self._draw_history_chart)

                except Exception as ex:
                    print("Erro no monitor:", ex)

                time.sleep(0.2)
        finally:
            pythoncom.CoUninitialize()

    # ---------- Configurações ----------
    def _open_settings_modal(self):
        def _cfg_preview_text(tmp_cfg):
            # exemplos práticos a partir do perfil diário
            t85 = allowed_time_seconds_for_level(85, tmp_cfg)
            t90 = allowed_time_seconds_for_level(90, tmp_cfg)
            def pretty(sec):
                s = int(max(0, sec)); h = s // 3600; m = (s % 3600) // 60
                if h > 0: return f"{h}h {m}min"
                return f"{m}min"
            return (f"Exemplo prático (base diária 8h):\n"
                    f"• A 85 dB: ~{pretty(t85)} até atingir 100% da dose diária.\n"
                    f"• A 90 dB: ~{pretty(t90)} até atingir 100% da dose diária.\n"
                    f"O modo Prefixado mantém uma folga mínima antes de ajustar o volume.")

        def _make_tmp_cfg(ref_db=None, er=None, min_vol=None, def_vol=None):
            tmp = dict(self.cfg)
            if ref_db is not None: tmp["ref_db"] = float(ref_db)
            if er is not None:     tmp["exchange_rate_db"] = float(er)
            # base diária é sempre 8h
            tmp["base_time_sec"] = 8 * 3600.0
            if min_vol is not None: tmp["min_enforced_volume"] = float(min_vol)
            if def_vol is not None: tmp["default_volume"] = float(def_vol)
            return tmp

        PROFILES = {
            "NIOSH (85 dB / 8h, 3 dB)": {
                "ref_db": 85.0, "er": 3.0,
                "desc": "Padrão ocupacional: 85 dB por 8h, troca 3 dB."
            },
            "OMS (80 dB / 8h, 3 dB)": {
                "ref_db": 80.0, "er": 3.0,
                "desc": "Mais protetivo: 80 dB por 8h, troca 3 dB."
            },
            "Personalizado (manter atual)": {
                "ref_db": None, "er": None,
                "desc": "Mantém valores atuais de referência e troca. Base diária é sempre 8h."
            },
        }

        top = ctk.CTkToplevel(self)
        top.title("Configurações")
        top.geometry("700x720")
        top.attributes("-topmost", True)

        tab = ctk.CTkTabview(top, width=660, height=590)
        tab.pack(fill="both", expand=True, padx=16, pady=16)
        basic = tab.add("Básico")
        adv = tab.add("Avançado")

        # ===== Básico =====
        basic_wrap = ctk.CTkFrame(basic, fg_color=DISCORD_SURFACE)
        basic_wrap.pack(fill="both", expand=True, padx=6, pady=6)

        ctk.CTkLabel(basic_wrap, text="Nível de proteção (perfil diário):", anchor="w").pack(fill="x", pady=(8,4))
        profile_names = list(PROFILES.keys())

        # Detecta se atual casa com NIOSH ou OMS
        cur_ref = round(float(self.cfg["ref_db"]), 1)
        cur_er  = round(float(self.cfg["exchange_rate_db"]), 1)
        current_profile = "Personalizado (manter atual)"
        if abs(cur_ref - 85.0) < 0.6 and abs(cur_er - 3.0) < 0.6:
            current_profile = "NIOSH (85 dB / 8h, 3 dB)"
        elif abs(cur_ref - 80.0) < 0.6 and abs(cur_er - 3.0) < 0.6:
            current_profile = "OMS (80 dB / 8h, 3 dB)"

        var_profile = tk.StringVar(value=current_profile)
        opt_profile = ctk.CTkOptionMenu(basic_wrap, values=profile_names, variable=var_profile)
        opt_profile.pack(fill="x", pady=(0,6))

        lbl_profile_desc = ctk.CTkLabel(basic_wrap, text=PROFILES[current_profile]["desc"],
                                        text_color="#B5BAC1", justify="left", wraplength=600)
        lbl_profile_desc.pack(fill="x", pady=(0,10))

        # Estratégia do Dinâmico
        ctk.CTkLabel(basic_wrap, text="Estratégia do Dinâmico:", anchor="w").pack(fill="x", pady=(6,4))
        dyn_names = ["Reserva de tempo (10–20 min)", "Reduzir até Zona Segura"]
        dyn_map_name_to_key = {
            "Reserva de tempo (10–20 min)": "reserva",
            "Reduzir até Zona Segura": "zona_segura",
        }
        dyn_map_key_to_name = {v: k for k, v in dyn_map_name_to_key.items()}
        var_dyn = tk.StringVar(value=dyn_map_key_to_name.get(self.dynamic_strategy, dyn_names[0]))
        ctk.CTkOptionMenu(basic_wrap, values=dyn_names, variable=var_dyn).pack(fill="x", pady=(0,6))
        ctk.CTkLabel(basic_wrap,
            text=("• Reserva: mantém uma folga alvo e reduz suave quando precisa.\n"
                  "• Zona Segura: baixa o volume gradualmente até o ponteiro ficar no verde."),
            text_color="#B5BAC1", justify="left", wraplength=600).pack(fill="x", pady=(0,10))

        def _mk_slider_row(parent, title, init_val, on_change, max_to=60):
            row = ctk.CTkFrame(parent, fg_color=DISCORD_SURFACE)
            row.pack(fill="x", pady=6)
            ctk.CTkLabel(row, text=title, anchor="w").pack(side="left", padx=(0, 8))
            init_val2 = max(0.0, min(float(init_val), float(max_to)))
            val_lbl = ctk.CTkLabel(row, text=f"{int(round(init_val2))}%", width=48, anchor="e")
            val_lbl.pack(side="right")
            sld = ctk.CTkSlider(
                row, from_=0, to=max_to, number_of_steps=int(max_to), width=380,
                command=lambda v:(val_lbl.configure(text=f"{int(float(v))}%"), on_change(float(v))))
            sld.set(init_val2); sld.pack(fill="x", padx=(0,56))
            return sld, val_lbl

        tmp_min = [self.cfg["min_enforced_volume"]]
        tmp_def = [self.cfg["default_volume"]]
        ctk.CTkLabel(basic_wrap, text="Volumes:", anchor="w").pack(fill="x", pady=(8,0))
        sld_def,_ = _mk_slider_row(basic_wrap, "Volume inicial ao abrir o app",
                                   float(self.cfg["default_volume"]),
                                   lambda v: tmp_def.__setitem__(0, v), max_to=60)
        sld_min,_ = _mk_slider_row(basic_wrap, "Volume mínimo imposto em bloqueios",
                                   float(self.cfg["min_enforced_volume"]),
                                   lambda v: tmp_min.__setitem__(0, v), max_to=60)

        # Soft-lock
        var_dyn_softlock = tk.BooleanVar(value=self.dynamic_softlock_enabled)
        ctk.CTkCheckBox(
            basic_wrap,
            text="Travar aumentos enquanto o Dinâmico reduz (soft-lock)",
            variable=var_dyn_softlock
        ).pack(anchor="w", pady=4)

        preview_box = ctk.CTkFrame(basic_wrap, fg_color=DISCORD_SURFACE_ALT, corner_radius=8)
        preview_box.pack(fill="x", pady=(8,4))
        lbl_preview = ctk.CTkLabel(preview_box, text=_cfg_preview_text(self.cfg), justify="left", wraplength=600)
        lbl_preview.pack(padx=12, pady=10)

        def _refresh_preview_for_profile(name):
            p = PROFILES.get(name, PROFILES["Personalizado (manter atual)"])
            if p["ref_db"] is None:
                tmp = _make_tmp_cfg(min_vol=tmp_min[0], def_vol=tmp_def[0])
            else:
                tmp = _make_tmp_cfg(ref_db=p["ref_db"], er=p["er"], min_vol=tmp_min[0], def_vol=tmp_def[0])
            lbl_profile_desc.configure(text=p["desc"])
            lbl_preview.configure(text=_cfg_preview_text(tmp))
        opt_profile.configure(command=_refresh_preview_for_profile)

        # ===== Avançado =====
        adv_wrap = ctk.CTkFrame(adv, fg_color=DISCORD_SURFACE)
        adv_wrap.pack(fill="both", expand=True, padx=6, pady=6)

        ctk.CTkLabel(adv_wrap, text="(Avançado) Ajustes técnicos",
                     text_color="#B5BAC1").pack(anchor="w", pady=(0,8))

        def add_row(parent, label, initial, placeholder="", width=140):
            row = ctk.CTkFrame(parent, fg_color=DISCORD_SURFACE)
            row.pack(fill="x", pady=6)
            ctk.CTkLabel(row, text=label, width=320, anchor="w").pack(side="left")
            entry = ctk.CTkEntry(row, width=width, placeholder_text=placeholder)
            entry.pack(side="right"); entry.insert(0, str(initial))
            return entry

        e_min_db = add_row(adv_wrap, "Mínimo dB (escala do gauge):", self.cfg["min_db"])
        e_max_db = add_row(adv_wrap, "Máximo dB (escala do gauge):", self.cfg["max_db"])
        e_ref_db = add_row(adv_wrap, "Nível de referência (dB):", self.cfg["ref_db"])
        e_er     = add_row(adv_wrap, "Taxa de troca (dB) [3]:", self.cfg["exchange_rate_db"])

        btns = ctk.CTkFrame(top, fg_color=DISCORD_SURFACE); btns.pack(fill="x", pady=(0,12), padx=16)

        def apply_all():
            try:
                # Perfil (diário 8h)
                chosen = var_profile.get()
                if chosen != "Personalizado (manter atual)":
                    p = PROFILES[chosen]
                    ref_db = float(p["ref_db"]); er = float(p["er"])
                else:
                    ref_db = float(e_ref_db.get()); er = float(e_er.get())

                min_db = float(e_min_db.get()); max_db = float(e_max_db.get())
                if max_db - min_db < 10.0: raise ValueError("Max dB deve ser ≥10 acima do Min dB.")
                if er <= 0: raise ValueError("Taxa de troca (dB) deve ser > 0.")

                min_vol = float(tmp_min[0]); def_vol = float(tmp_def[0])
                for v in (min_vol, def_vol):
                    if not (0.0 <= v <= 100.0):
                        raise ValueError("Volumes devem estar entre 0 e 100%.")

                # Estratégia do dinâmico
                dyn_key = {"Reserva de tempo (10–20 min)": "reserva", "Reduzir até Zona Segura": "zona_segura"}[var_dyn.get()]

                # Atualiza cfg diária (base fixa 8h)
                self.cfg.update({
                    "min_db": min_db,
                    "max_db": max_db,
                    "ref_db": ref_db,
                    "base_time_sec": 8 * 3600.0,
                    "exchange_rate_db": er,
                    "min_enforced_volume": min_vol,
                    "default_volume": def_vol,
                })
                self.dynamic_softlock_enabled = bool(var_dyn_softlock.get())
                self.dynamic_strategy = dyn_key

                # Aplica UI
                self._refresh_profile_label()
                self.gauge.set_bounds(self.cfg["min_db"], self.cfg["max_db"])
                self.gauge.set_profile_ref(self.cfg["ref_db"])
                L_eff = map_percent_to_db(self._vol_cache, self.cfg)
                self.gauge.set_value(L_eff, self.session_dose)
                self.vol_label.configure(text=f"{round_pct_ui(self._vol_cache)}%")
                self.set_mode(self.mode, silent=True)
                self._save_settings()

                _refresh_preview_for_profile(var_profile.get())
                messagebox.showinfo("Configurações", "Configurações aplicadas e salvas.")
            except Exception as ex:
                messagebox.showerror("Configurações", f"Erro: {ex}")

        def restore_defaults():
            var_profile.set("NIOSH (85 dB / 8h, 3 dB)")
            var_dyn.set("Reserva de tempo (10–20 min)")
            var_dyn_softlock.set(True)
            sld_def.set(self._defaults_cfg["default_volume"])
            sld_min.set(self._defaults_cfg["min_enforced_volume"])
            e_min_db.delete(0, tk.END); e_min_db.insert(0, str(self._defaults_cfg["min_db"]))
            e_max_db.delete(0, tk.END); e_max_db.insert(0, str(self._defaults_cfg["max_db"]))
            e_ref_db.delete(0, tk.END); e_ref_db.insert(0, "85")
            e_er.delete(0, tk.END); e_er.insert(0, "3")
            _refresh_preview_for_profile("NIOSH (85 dB / 8h, 3 dB)")
            messagebox.showinfo("Configurações", "Padrões restaurados (não esqueça de clicar em Aplicar).")

        ctk.CTkButton(btns, text="Aplicar", fg_color=DISCORD_ACCENT, width=120, command=apply_all)\
            .pack(side="left", padx=6)
        ctk.CTkButton(btns, text="Restaurar padrões", fg_color="#6b7280", width=160, command=restore_defaults)\
            .pack(side="left", padx=6)
        ctk.CTkButton(btns, text="Fechar", fg_color="#444", width=120, command=top.destroy)\
            .pack(side="right", padx=6)

        _refresh_preview_for_profile(var_profile.get())

    # ----------
    def _on_close(self):
        self._stop_event.set()
        try: self._save_settings()
        except Exception: pass
        try: self._stop_lock_enforcer()
        except Exception: pass
        try:
            if self._ui_com_inited: pythoncom.CoUninitialize()
        except Exception: pass
        self.destroy()


if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    app = SoundMonitorApp()
    app.mainloop()
