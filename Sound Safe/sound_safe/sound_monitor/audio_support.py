# audio_support.py

import platform

_PYCAW_AVAILABLE = False
try:
    if platform.system() == "Windows":
        from comtypes import CLSCTX_ALL  # type: ignore
        from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume  # type: ignore
        _PYCAW_AVAILABLE = True
    else:
        CLSCTX_ALL = None
        AudioUtilities = None
        IAudioEndpointVolume = None
except Exception:
    _PYCAW_AVAILABLE = False
    CLSCTX_ALL = None
    AudioUtilities = None
    IAudioEndpointVolume = None
