# Исправление Запуск_на_Windows.bat

## Проблема
Старый `Запуск_на_Windows.bat` был сохранён как UTF-8 **без BOM** с кириллицей (`Администраций` и т.д.).
Windows CMD по умолчанию работает в OEM 866, а не UTF-8 — поэтому 2-байтовые UTF-8 символы ломали парсер:
- `if errorlevel 1 (` превращалось в `if errorlevel` + мусор → ошибка `'errorlevel' is not recognized`
- `echo Если окно не открылось` резалось на куски → `'крылось' is not recognized`, `'vel'`, `'лась'` и т.д.

## Исправление (текущая версия)
- BAT пересохранён **UTF-8 с BOM (EF BB BF)** + `CRLF` — CMD с `chcp 65001` теперь корректно читает файл
- **Вся кириллица убрана из BAT** — только ASCII (English) → 100% совместимость с любой локалью Windows
- Добавлен `title DaData Parser - Administrations`
- Исправлен `if %errorlevel% neq 0 (` вместо `if errorlevel 1 (` — более надёжно
- Добавлена копия GUI с ASCII-именем: `DaData_GUI.py` (оригинал `DaData_Администрации_GUI.py` тоже остался)
- BAT теперь запускает `python DaData_GUI.py` — без кириллицы в имени файла, чтобы не зависеть от codepage пути
- Добавлен `Запуск_PowerShell.ps1` — альтернатива для PowerShell (лучше работает с UTF-8, просто правой кнопкой → Запуск в PowerShell)

## Как запускать
1. **Двойной клик** по `Запуск_на_Windows.bat` — для CMD
2. Или правой кнопкой по `Запуск_PowerShell.ps1` → *Запуск с помощью PowerShell*
3. Или вручную: `pip install openpyxl requests` → `python DaData_GUI.py`

Если снова увидите `errorlevel` — убедитесь, что запускаете именно новый файл из ветки `arena/363bbf0c-gov` (размер нового BAT 1094 байт, старый был 1315).
