# Dadata Parser — отдельный проект

Парсер ЕГРЮЛ по ИНН через **DaData Find-Party** (`findById/party`) с:
- **SOCKS5 прокси** `IP:PORT:LOGIN:PASS` (авторотация провайдера — менять не нужно, но поддерживается пул)
- **Ротацией токенов**: до **10 000 запросов/сутки на 1 токен**, после — автопереключение на следующий
- **Учётом расхода по токенам в сутки** — лог + SQLite + консоль
- Настраиваемым пулом токенов и прокси через `config.yaml`
- Resume по `JSONL`, Excel-обогащением и CLI/GUI

> Выделен из `scripts/enrich_dadata.py` + `DaData_GUI.py` ветки `arena/363bbf0c-gov`.

---

## Структура

```
dadata_parser/
  dadata_parser/
    __init__.py
    __main__.py        # CLI entry
    config.py          # загрузка config.yaml / env
    token_manager.py   # учёт 10к/сутки, ротация
    proxy_manager.py   # SOCKS5 IP:PORT:LOGIN:PASS → socks5:// URL
    dadata_client.py   # requests + SOCKS5 + ретраи
    storage.py         # SQLite daily counters + JSONL
    worker.py          # оркестратор (resume, прогресс, Excel)
    gui/
      app.py           # Tkinter GUI с настройками токенов/прокси
  config.yaml.example  # пример настроек
  config.yaml          # ваш локальный конфиг (gitignored)
  requirements.txt
  README.md
```

---

## Установка

```bash
cd dadata_parser
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
cp config.yaml.example config.yaml
# отредактируйте config.yaml — вставьте токены и прокси
```

---

## Настройка `config.yaml`

```yaml
dadata:
  tokens:
    - "77eff8edd3640fc0002ad6ae2431261a2f382dfb"
    - "второй_токен..."
    - "третий_токен..."
  daily_limit_per_token: 10000    # лимит DaData на ключ в сутки
  timeout: 15
  retries: 3
  delay_between_requests: 0.2     # 0.2 = 5 req/s на прокси

proxy:
  enabled: true
  # Формат: IP:PORT:LOGIN:PASS  (колонка-разделитель, 4 части)
  # Также поддерживается IP:PORT;LOGIN:PASS
  # Пул может быть 1 прокси с авторотацией — не обязательно много
  list:
    - "95.78.126.13:14755:5nqb8nr:9cz93bt"
  rotation: "round_robin"        # round_robin | random | sticky
  verify_ssl: true

storage:
  sqlite_path: "./data/parser.db"
  raw_jsonl: "./data/dadata_raw.jsonl"
  log_path: "./data/parser.log"

output:
  excel_template: null            # null = без обогащения Excel, или путь к входному xlsx
  excel_output: "./data/enriched.xlsx"
  save_raw_jsonl: true
```

Также токены можно передать через ENV:
```bash
DADATA_TOKENS="token1,token2,token3" python -m dadata_parser --input inn_list.txt
```

---

## Формат прокси

Поддерживаются строки:

```
95.78.126.13:14755:5nqb8nr:9cz93bt      # IP:PORT:LOGIN:PASS
95.78.126.13:14755;5nqb8nr;9cz93bt      # IP:PORT;LOGIN;PASS (альтернатива)
95.78.126.13:14755                      # без авторизации (редко)
```

Конвертируется в `socks5h://LOGIN:PASS@IP:PORT` (через `PySocks` + `requests[socks]`).
Авторотация провайдера = один прокси в списке достаточно, он сам меняет IP.

Проверка прокси:
```bash
python -m dadata_parser --test-proxy
```

---

## Логика ротации токенов

1. Каждый запрос требует токен.
2. `TokenManager` хранит в SQLite `usage(token, date) -> count`.
3. Перед запросом берёт первый токен у которого `count < daily_limit_per_token` (10к).
4. После успешных запросов инкремент.
5. При `429` или исчерпании — помечает токен на сегодня как exhausted, переключает на следующий.
6. Если все токены исчерпаны — `AllTokensExhausted` → остановка с сообщением «лимит на сутки исчерпан, следующий слот в 00:00 МСК».

### Лог расхода

- Консоль (каждые 50 запросов + прогресс-бар):
  ```
  [12:34:56] token 77eff8...f382dfb: 420/10000 (4.2%) | токен2: 0/10000 | done 420/1500
  ```
- Файл `data/parser.log` (ротация).
- SQLite `data/parser.db` таблица `token_usage(date, token_masked, count, limit)` — можно строить графики.
- JSONL `data/dadata_raw.jsonl` — построчно `{inn, queried_at, token_masked, response}` для resume.

Пример SQL:
```sql
SELECT date, substr(token,1,6)||'...' as token, count, limit FROM token_usage WHERE date='2026-11-16';
```

---

## CLI

```bash
# 1. Парс списка ИНН из txt/csv/xlsx
python -m dadata_parser --input data/inns.txt --output data/out.jsonl

# 2. Обогащение Excel (как в старом enrich_dadata.py) — добавляет 15 зелёных колонок
python -m dadata_parser --input Администраций_городов_из_реестра_2026-11-10.xlsx --excel-output result.xlsx

# 3. Тест 1 ИНН (Пермь)
python -m dadata_parser --test --inn 5902290635

# 4. Проверка прокси + токена
python -m dadata_parser --test-proxy --test

# 5. Указать токены через CLI
python -m dadata_parser --tokens token1,token2 --proxy 95.78.126.13:14755:5nqb8nr:9cz93bt --input inns.txt

# 6. GUI
python -m dadata_parser --gui
```

Опции `--help` покажет все флаги.

---

## GUI

```bash
python -m dadata_parser --gui
# или
python -m dadata_parser.gui.app
```

GUI позволяет:
- Редактировать таблицу токенов (добавить/удалить, видеть остаток на сегодня)
- Редактировать список прокси
- Выбрать входной файл, задержку, resume
- Видеть live-лог расхода по токенам
- Тест прокси/токена в 1 клик

---

## API: DaData Find-Party

```
POST https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/party
Headers: Authorization: Token <API_KEY>
Body: {"query": "7707083893"}
```

Ответ → `suggestions[0].data` содержит `management.name`, `management.post`, `state.status`, `ogrn`, `address`, `okved`, `emails`, `phones` и т.д. Сохраняется сырой JSONL для дозабора любых полей без повторных запросов.

---

## Лимиты и рекомендации

- **10 000 / сутки / токен** — считать по Москве (DaData бьёт по календарным суткам).
- `delay_between_requests: 0.2` — защита от 429 (5 r/s на прокси достаточно; при 429 — экспоненциальная пауза).
- Resume: если прервали — повторный запуск продолжает с места (`JSONL` + `SQLite`).
- Прокси с авторотацией — 1 строки достаточно; при пуле — `round_robin`.

---

## Разработка

```bash
pip install -r requirements-dev.txt  # если есть
pytest
```

Логи: `data/parser.log`, `data/token_usage_2026-11-16.log`

---

## Лицензия

MIT — выделили из внутреннего проекта Arena GIS.
