@echo off
chcp 65001 >nul
title Пересобрать Excel из JSONL
python "%~dp0пересобрать_excel.py"
if %errorlevel% neq 0 (
  py "%~dp0пересобрать_excel.py"
)
pause
