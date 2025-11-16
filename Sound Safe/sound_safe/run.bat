@echo off
setlocal
cd /d "%~dp0"
python -m sound_monitor.main || python run.py
