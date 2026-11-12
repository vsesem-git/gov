@echo off
chcp 65001 >nul
title DaData Парсер Администраций — Windows
echo ===============================================
echo  DaData Парсер Администраций городов
echo  Ветка arena/363bbf0c-gov
echo ===============================================
echo.

:: Проверка Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ОШИБКА] Python не найден! Установите Python 3.10+ с https://python.org
    echo Не забудьте поставить галочку "Add Python to PATH" при установке
    pause
    exit /b 1
)

echo [1/3] Проверка зависимостей...
pip install openpyxl requests --quiet
if errorlevel 1 (
    echo [ОШИБКА] Не удалось установить зависимости
    pause
    exit /b 1
)

echo [2/3] Запуск GUI...
echo Если окно не открылось — запустите вручную: python DaData_Администрации_GUI.py
echo.

:: Запуск GUI
python DaData_Администрации_GUI.py

if errorlevel 1 (
    echo.
    echo [ОШИБКА] Программа завершилась с ошибкой
    pause
)

echo.
echo [3/3] Готово. Результаты в папке с программой.
pause
