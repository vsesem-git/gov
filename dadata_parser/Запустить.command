#!/bin/bash
# macOS / Linux double-click launcher
cd "$(dirname "$0")"
python3 Запуск.py
if [ $? -ne 0 ]; then python Запуск.py; fi
