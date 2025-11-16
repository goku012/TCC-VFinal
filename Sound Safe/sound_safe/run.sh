#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"
python -m sound_monitor.main || python3 run.py
