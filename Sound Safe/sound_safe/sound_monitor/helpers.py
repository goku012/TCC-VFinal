# helpers.py

import math

def map_percent_to_db(vol_percent, cfg):
    min_db = cfg["min_db"]; max_db = cfg["max_db"]
    v = max(0, min(100, float(vol_percent)))
    return min_db + (max_db - min_db) * (v / 100.0)

def db_to_percent(L, cfg):
    min_db = cfg["min_db"]; max_db = cfg["max_db"]
    ratio = (float(L) - min_db) / max(1e-9, (max_db - min_db))
    return max(0.0, min(100.0, ratio * 100.0))

def allowed_time_seconds_for_level(L, cfg):
    """
    CÁLCULO DIÁRIO: base SEMPRE 8 horas no nível de referência do perfil (OMS=80/8h ou NIOSH=85/8h).
    Troca = 3 dB.
    """
    base_8h = float(cfg["base_time_sec"])               # 8h
    ref = float(cfg["ref_db"])                          # 80 (OMS) ou 85 (NIOSH)
    er = max(0.1, float(cfg.get("exchange_rate_db", 3.0)))  # 3 dB
    diff = float(L) - ref
    allowed = base_8h * (2 ** (-diff / er))
    return max(1.0, allowed)

def dose_increment_per_second(L, cfg):
    return 1.0 / allowed_time_seconds_for_level(L, cfg)

def risk_zone_from_dose(dose):
    pct = dose * 100.0
    if pct < 50.0:
        return "SEGURA"
    elif pct < 100.0:
        return "ATENÇÃO"
    else:
        return "PERIGO"

def risk_zone_from_level(L, cfg):
    safe_cut = float(cfg["ref_db"]) - 15.0
    warn_cut = float(cfg["ref_db"])
    if L < safe_cut:
        return "SEGURA"
    elif L < warn_cut:
        return "ATENÇÃO"
    else:
        return "PERIGO"

def fmt_hms(seconds):
    s = int(max(0, seconds)); h = s // 3600; m = (s % 3600) // 60; sec = s % 60
    return f"{h:02d}:{m:02d}:{sec:02d}"

def round_pct_ui(x: float) -> int:
    return int(math.floor(float(x) + 0.5))
