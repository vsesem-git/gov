#!/usr/bin/env python3
"""
Обогащение реестра администраций городов данными DaData Find-Party по ИНН.

API: https://dadata.ru/api/find-party/
Endpoint: POST https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/party
Headers:
  Content-Type: application/json
  Accept: application/json
  Authorization: Token <API_KEY>

Что делает скрипт:
- читает Администраций_городов_из_реестра_2026-11-10.xlsx (все 4 листа)
- для каждого уникального ИНН делает запрос к DaData
- сохраняет сырой ответ каждого запроса в JSONL (для последующего дозабора любых полей)
- сверяет email/сайт из реестра с данными DaData
- дописывает в таблицу столбцы: Руководитель ФИО, Должность, Email/Site/Телефон из DaData, флаги совпадения
- сохраняет новую книгу + отдельный JSON-дамп

Запуск:
  DADATA_TOKEN=77eff8edd3640fc0002ad6ae2431261a2f382dfb python3 enrich_dadata.py
  # или:
  python3 enrich_dadata.py --token 77eff8edd3640fc0002ad6ae2431261a2f382dfb --input "Администрации_городов_из_реестра_2026-11-10.xlsx"

Требует: pip install openpyxl requests
"""
import argparse, json, time, os, sys, pathlib
from collections import defaultdict
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import requests

DADATA_URL = "https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/party"
DEFAULT_INPUT = "Администрации_городов_из_реестра_2026-11-10.xlsx"
DEFAULT_OUTPUT = "Администрации_городов_обогащенная_DaData.xlsx"
DEFAULT_RAW = "dadata_raw_responses.jsonl"
DEFAULT_TOKEN = "77eff8edd3640fc0002ad6ae2431261a2f382dfb"

# Новые столбцы, которые будут добавлены справа от существующих 14
NEW_COLUMNS = [
    "Руководитель ФИО (DaData)",
    "Должность руководителя (DaData)",
    "Тип руководителя (DaData)",
    "Email из ЕГРЮЛ (DaData)",
    "Сайт из ЕГРЮЛ (DaData)",
    "Телефон из ЕГРЮЛ (DaData)",
    "Email совпадает с реестром?",
    "Сайт совпадает с реестром?",
    "Дополнительный Email найден?",
    "Дополнительный сайт найден?",
    "ОГРН из DaData",
    "Наименование полное из ЕГРЮЛ (DaData)",
    "Статус организации (DaData)",
    "Дата актуальности ЕГРЮЛ (DaData)",
    "Статус запроса DaData",
    "Дата запроса DaData",
    "Сырой ответ (сжатый JSON)",
]

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=10, name="Calibri")
NEW_HEADER_FILL = PatternFill(start_color="548235", end_color="548235", fill_type="solid") # зеленый для новых
THIN_BORDER = Border(left=Side(style='thin', color='B4C6E7'), right=Side(style='thin', color='B4C6E7'), top=Side(style='thin', color='B4C6E7'), bottom=Side(style='thin', color='B4C6E7'))
ALT_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

def normalize_str(s):
    if s is None: return ""
    return str(s).strip()

def normalize_site(site):
    if not site: return ""
    s = str(site).strip().lower()
    # убрать протокол, www., слэш в конце
    for p in ["https://","http://","www."]:
        if s.startswith(p):
            s = s[len(p):]
    s = s.rstrip("/").strip()
    return s

def normalize_email(email):
    if not email: return ""
    return str(email).strip().lower()

def call_dadata(inn, token, session, retries=3):
    payload = {"query": inn}
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Authorization": f"Token {token}",
    }
    last_err = None
    for attempt in range(retries):
        try:
            resp = session.post(DADATA_URL, json=payload, headers=headers, timeout=15)
            if resp.status_code == 429:
                wait = 2 ** attempt + 1
                print(f"[429] Too Many Requests для ИНН {inn}, пауза {wait}с (попытка {attempt+1})")
                time.sleep(wait)
                continue
            if resp.status_code != 200:
                return {"inn": inn, "http_status": resp.status_code, "error": resp.text[:2000], "attempt": attempt+1}
            data = resp.json()
            return {"inn": inn, "http_status": 200, "data": data}
        except requests.exceptions.RequestException as e:
            last_err = str(e)
            wait = 1.5 ** attempt
            print(f"[ERR] {inn}: {e}, retry {attempt+1} через {wait:.1f}с")
            time.sleep(wait)
    return {"inn": inn, "http_status": None, "error": last_err or "max retries", "data": None}

def extract_fields(dadata_resp):
    """Вытягивает нужные поля из ответа DaData findById/party"""
    # Ответ: {"suggestions":[{"value":"...","unrestricted_value":"...","data":{...}}]}
    if not dadata_resp or dadata_resp.get("http_status") != 200 or not dadata_resp.get("data"):
        return {
            "fio": "",
            "post": "",
            "management_type": "",
            "email": "",
            "site": "",
            "phone": "",
            "ogrn": "",
            "full_name": "",
            "status": "",
            "actuality": "",
            "raw_compact": json.dumps(dadata_resp, ensure_ascii=False)[:32000] if dadata_resp else ""
        }
    data_wrapper = dadata_resp["data"]
    sugg = data_wrapper.get("suggestions") or []
    if not sugg:
        return {
            "fio": "", "post": "", "management_type": "", "email": "", "site": "", "phone": "",
            "ogrn": "", "full_name": "", "status": "Не найдено в ЕГРЮЛ", "actuality": "",
            "raw_compact": json.dumps(dadata_resp, ensure_ascii=False)[:32000]
        }
    d = sugg[0].get("data") or {}
    # Руководитель: d.management.name, d.management.post
    mgmt = d.get("management") or {}
    fio = normalize_str(mgmt.get("name"))
    post = normalize_str(mgmt.get("post"))
    mtype = normalize_str(mgmt.get("type")) # e.g. EMPLOYEE, etc
    # Email/phone/site могут быть на разных уровнях: d.email, d.site, d.phone ; также иногда в контактах
    # DaData в findById иногда не отдаёт email/site — они доступны в ЕГРЮЛ через другие методы,
    # но попробуем вытащить все что есть
    email = normalize_str(d.get("email"))
    site = normalize_str(d.get("site"))
    phone = normalize_str(d.get("phone"))
    # ОГРН
    ogrn = normalize_str(d.get("ogrn"))
    full_name = normalize_str(sugg[0].get("unrestricted_value") or sugg[0].get("value") or d.get("name",{}).get("full_with_opf") or d.get("name",{}).get("full"))
    state = d.get("state") or {}
    status = normalize_str(state.get("status")) # ACTIVE, LIQUIDATING, LIQUIDATED
    actuality = normalize_str(state.get("actuality_date")) or normalize_str(state.get("registration_date"))
    # Дополнительные email/сайты могут быть в массиве? Проверяем
    # Иногда DaData кладет контакты в d.contacts etc — сохраним сырой JSON для дозабора

    return {
        "fio": fio,
        "post": post,
        "management_type": mtype,
        "email": email,
        "site": site,
        "phone": phone,
        "ogrn": ogrn,
        "full_name": full_name,
        "status": status,
        "actuality": actuality,
        "raw_compact": json.dumps(dadata_resp, ensure_ascii=False)[:32000]
    }


def style_sheet(ws, new_col_start):
    # Автофильтр, заморозка, ширины
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = '1:1'
    # Ширины
    widths = [55,30,30,32,16,14,13,32,18,26,22,32,16,18, 26,28,18,26,26,18,18,18,18,16,26,28,16,16,32]
    for i, w in enumerate(widths, start=1):
        if i <= ws.max_column:
            ws.column_dimensions[get_column_letter(i)].width = w
    # Стили заголовка
    for cell in ws[1]:
        col_idx = cell.column
        if col_idx >= new_col_start:
            cell.fill = NEW_HEADER_FILL
        else:
            cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = THIN_BORDER
    ws.row_dimensions[1].height = 38
    # Строки
    for r_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        is_alt = (r_idx % 2 == 0)
        for cell in row:
            cell.alignment = Alignment(vertical='center', wrap_text=True)
            cell.border = THIN_BORDER
            cell.font = Font(size=9, name="Calibri")
            if is_alt:
                cell.fill = ALT_FILL
        ws.row_dimensions[r_idx].height = 16

def enrich_file(input_path, output_path, raw_path, token, delay=0.2, resume=True):
    print(f"Чтение {input_path} ...")
    wb = openpyxl.load_workbook(input_path, read_only=False, data_only=True)
    # Собираем все ИНН по всем листам кроме Оглавления
    inn_to_rows = defaultdict(list) # inn -> list of (sheet, row_idx, row_data)
    inn_set = set()
    sheets_to_process = [ws for ws in wb.worksheets if ws.title != "Оглавление"]
    for ws in sheets_to_process:
        rows = list(ws.iter_rows(values_only=True))
        header = rows[0]
        # Найти индекс ИНН (колонка F = 6)
        try:
            inn_idx = header.index("ИНН")
        except ValueError:
            inn_idx = 5
        for r_idx, row in enumerate(rows[1:], start=2):
            inn = normalize_str(row[inn_idx])
            if inn and inn.isdigit() and len(inn) in (10,12):
                inn_set.add(inn)
                inn_to_rows[inn].append((ws.title, r_idx))
    print(f"Найдено уникальных ИНН: {len(inn_set)} (из {sum(len(list(ws.iter_rows()))-1 for ws in sheets_to_process)} строк)")
    # Загрузка уже сохранённых ответов (resume)
    existing = {}
    if resume and os.path.exists(raw_path):
        print(f"Найден существующий {raw_path}, подгружаем для resume...")
        with open(raw_path, "r", encoding="utf-8") as f:
            for line in f:
                try:
                    obj = json.loads(line)
                    inn = obj.get("inn")
                    if inn:
                        existing[inn] = obj.get("response") or obj
                except: continue
        print(f"  загружено {len(existing)} ответов")
    to_query = [inn for inn in inn_set if inn not in existing]
    print(f"К запросу осталось: {len(to_query)} ИНН")
    # Подготовка сессии
    session = requests.Session()
    session.headers.update({"User-Agent": "Arena-GIS-enricher/1.0"})
    # Файл для дозаписи raw
    raw_file = open(raw_path, "a", encoding="utf-8") if resume else open(raw_path, "w", encoding="utf-8")
    # Если файл новый и пустой, ничего не делаем; если resume и файл существует, мы уже не перезаписываем
    results = {}
    # Заполним results из existing (для тех кто уже есть)
    for inn, resp in existing.items():
        # existing хранит уже обёрнутую структуру? Нормализуем
        if "http_status" in resp:
            results[inn] = resp
        else:
            # старый формат где хранился просто ответ DaData
            results[inn] = resp

    # Основной цикл запросов
    start_time = time.time()
    for idx, inn in enumerate(sorted(to_query), start=1):
        print(f"[{idx}/{len(to_query)}] Запрос ИНН {inn} ...", end=" ", flush=True)
        resp = call_dadata(inn, token, session)
        results[inn] = resp
        # Сохраняем сырой ответ построчно (JSONL) — удобно для дальнейшего парсинга любых полей
        record = {
            "inn": inn,
            "queried_at": datetime.now().isoformat(),
            "response": resp
        }
        raw_file.write(json.dumps(record, ensure_ascii=False) + "\n")
        raw_file.flush()
        # Прогресс каждые 50
        if idx % 50 == 0:
            elapsed = time.time() - start_time
            rps = idx/elapsed if elapsed>0 else 0
            print(f"\n  прогресс {idx}/{len(to_query)}, {rps:.1f} req/s, elapsed {elapsed/60:.1f} мин")
        else:
            # краткий вывод статуса
            status = resp.get("http_status")
            has_data = "найдено" if resp.get("data") and resp["data"].get("suggestions") else "пусто" if status==200 else f"ошибка {status}"
            print(f"-> {has_data}")
        time.sleep(delay)
        # Бэкап каждые 200
        if idx % 200 == 0:
            raw_file.flush()
            os.fsync(raw_file.fileno())
    raw_file.close()
    # Теперь обогащаем каждую страницу
    print("\nОбогащение листов...")
    for ws in sheets_to_process:
        print(f"  Лист {ws.title} ...")
        rows = list(ws.iter_rows(values_only=True))
        header = list(rows[0])
        new_col_start = len(header) + 1
        # Добавляем новые заголовки
        for col_name in NEW_COLUMNS:
            header.append(col_name)
        # Перезаписываем заголовок в sheet
        ws.delete_rows(1)
        ws.insert_rows(1)
        for c_idx, val in enumerate(header, start=1):
            cell = ws.cell(row=1, column=c_idx, value=val)
        # Для каждой строки данных добавляем обогащение
        # Найти индексы исходных столбцов
        # Важно: после добавления новых столбцов старые индексы сохраняются (1..14)
        # Считываем заново строки (но мы уже удалили header, теперь данные начинаются со 2 строки)
        # Проще: пройтись по строкам и дописать в колонки new_col_start..
        # Текущий ws.max_row включает header + данные
        for r_idx in range(2, ws.max_row+1):
            inn = normalize_str(ws.cell(row=r_idx, column=6).value) # F = 6
            email_reg = normalize_str(ws.cell(row=r_idx, column=10).value) # J=10
            site_reg = normalize_str(ws.cell(row=r_idx, column=8).value) # H=8
            resp = results.get(inn) or {}
            fields = extract_fields(resp)
            # Подготовка значений
            fio = fields["fio"]
            post = fields["post"]
            mtype = fields["management_type"]
            email_dd = fields["email"]
            site_dd = fields["site"]
            phone_dd = fields["phone"]
            ogrn_dd = fields["ogrn"]
            full_dd = fields["full_name"]
            status_dd = fields["status"]
            actuality_dd = fields["actuality"]
            raw_compact = fields["raw_compact"]
            # Сравнение email/сайт
            email_match = "Да" if email_dd and normalize_email(email_dd)==normalize_email(email_reg) else ("Нет" if email_dd and email_reg else ("Пусто в DaData" if not email_dd else "Пусто в реестре" if not email_reg else "Нет"))
            site_match = "Да" if site_dd and normalize_site(site_dd)==normalize_site(site_reg) else ("Нет" if site_dd and site_reg else ("Пусто в DaData" if not site_dd else "Пусто в реестре" if not site_reg else "Нет"))
            has_extra_email = "Да" if email_dd and normalize_email(email_dd)!=normalize_email(email_reg) and email_dd else "Нет"
            has_extra_site = "Да" if site_dd and normalize_site(site_dd)!=normalize_site(site_reg) and site_dd else "Нет"
            # Статусы
            if resp.get("http_status")==200 and resp.get("data") and resp["data"].get("suggestions"):
                req_status = "Найдено"
            elif resp.get("http_status")==200:
                req_status = "Не найдено"
            elif resp.get("http_status"):
                req_status = f"Ошибка {resp.get('http_status')}"
            else:
                req_status = resp.get("error","Не запрошено")[:80]
            queried_at = datetime.now().strftime("%Y-%m-%d %H:%M")
            # Запись в ячейки
            values = [fio, post, mtype, email_dd, site_dd, phone_dd, email_match, site_match, has_extra_email, has_extra_site, ogrn_dd, full_dd, status_dd, actuality_dd, req_status, queried_at, raw_compact]
            for offset, val in enumerate(values, start=0):
                ws.cell(row=r_idx, column=new_col_start+offset, value=val)
        # Стилизация
        style_sheet(ws, new_col_start)
    # Сохраняем
    print(f"\nСохранение {output_path} ...")
    wb.save(output_path)
    print(f"Готово. Сырые ответы: {raw_path} ({os.path.getsize(raw_path)/1024:.1f} КБ)")
    print(f"Обогащённая книга: {output_path} ({os.path.getsize(output_path)/1024:.1f} КБ)")
    return output_path, raw_path

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Обогащение DaData по ИНН")
    parser.add_argument("--input", default=DEFAULT_INPUT, help="Входной Excel")
    parser.add_argument("--output", default=DEFAULT_OUTPUT, help="Выходной Excel")
    parser.add_argument("--raw", default=DEFAULT_RAW, help="Файл для сырых JSONL ответов")
    parser.add_argument("--token", default=os.getenv("DADATA_TOKEN", DEFAULT_TOKEN), help="DaData API ключ")
    parser.add_argument("--delay", type=float, default=0.2, help="Пауза между запросами, сек (0.2 = 5 req/s)")
    parser.add_argument("--resume", action="store_true", default=True, help="Продолжить с места остановки")
    parser.add_argument("--no-resume", dest="resume", action="store_false")
    args = parser.parse_args()
    if not args.token:
        print("Нужен токен DaData: --token или env DADATA_TOKEN")
        sys.exit(1)
    enrich_file(args.input, args.output, args.raw, args.token, delay=args.delay, resume=args.resume)
