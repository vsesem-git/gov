DaData Парсер Администраций — Windows
=====================================

1. Установите Python 3.10+ с https://python.org (галочка Add to PATH)
2. Дважды кликните Запуск_на_Windows.bat
   — он установит openpyxl/requests и откроет окно
3. В окне нажмите ▶ Запустить обогащение

Или вручную:
  pip install -r requirements.txt
  python DaData_Администрации_GUI.py

Токен DaData по умолчанию: 77eff8edd3640fc0002ad6ae2431261a2f382dfb
Меняйте в поле "DaData токен" при необходимости.

Результаты:
  - Администраци_городов_обогащенная_DaData.xlsx (зелёные колонки)
  - dadata_raw_responses.jsonl (каждая строка = 1 ИНН, для дозабора любых полей)

Подробно: README_WINDOWS_GUI.md
