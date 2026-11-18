@echo off
chcp 65001 >nul
title DaData Parser — запуск GUI
color 0B
echo ============================================================
echo  DaData Parser — простой запуск GUI
echo  Токены и прокси настраиваются ВНУТРИ окна — cmd не нужен
echo ============================================================
echo.

:: Проверка Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    py --version >nul 2>&1
    if %errorlevel% neq 0 (
        echo [ОШИБКА] Python не найден!
        echo Установите Python 3.10+ с https://python.org
        echo Обязательно отметьте "Add Python to PATH"
        echo.
        pause
        exit /b 1
    ) else (
        set PYTHON=py
    )
) else (
    set PYTHON=python
)

echo [1/2] Проверяю Python: 
%PYTHON% --version
echo.
echo [2/2] Запускаю GUI — установка зависимостей пройдет сама...
echo      (первый запуск может занять 30-60 сек)
echo.

:: Запускаем универсальный лаунчер
%PYTHON% "%~dp0Запуск.py"
if %errorlevel% neq 0 (
    echo.
    echo [ОШИБКА] Код %errorlevel%
    echo Попробуйте запустить вручную: %PYTHON% "%~dp0Запуск.py"
    pause
)
