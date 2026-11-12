#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DaData Парсер Администраций городов — Windows GUI
===================================================
Графический интерфейс для обогащения реестра администраций через DaData Find-Party.

Возможности:
- Выбор входного Excel (по умолчанию Администраций_городов_из_реестра_2026-11-10.xlsx)
- Ввод токена DaData (предзаполнен вашим ключом)
- Получение по ИНН: руководитель ФИО + должность, email/сайт/телефон из ЕГРЮЛ
- Проверка совпадения email/сайта с реестром (флаги Да/Нет)
- Сохранение КАЖДОГО сыроого ответа в JSONL (для дозабора любых полей позже)
- Прогресс-бар, лог, таблица предпросмотра, пауза/продолжение (resume)
- Итоговый Excel с 15 новыми зелёными колонками, фильтры + печать

Требования Windows:
  pip install openpyxl requests

Запуск:
  двойной клик по файлу или python DaData_Администрации_GUI.py
  или через Запуск_на_Windows.bat

Автор: Arena AI для ветки arena/363bbf0c-gov
API: https://dadata.ru/api/find-party/
Endpoint: POST https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/party
"""

import os
import sys
import json
import time
import threading
import queue
import webbrowser
from pathlib import Path
from datetime import datetime
from collections import defaultdict

# Попытка импорта обязательных библиотек с подсказкой
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("Нужно: pip install openpyxl")
    sys.exit(1)

try:
    import requests
except ImportError:
    print("Нужно: pip install requests")
    sys.exit(1)

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ===================== Константы =====================
DADATA_URL = "https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/party"
DEFAULT_TOKEN = "77eff8edd3640fc0002ad6ae2431261a2f382dfb"
DEFAULT_INPUT = "Администрации_городов_из_реестра_2026-11-10.xlsx"
DEFAULT_OUTPUT = "Администрации_городов_обогащенная_DaData.xlsx"
DEFAULT_RAW = "dadata_raw_responses.jsonl"

NEW_COLUMNS = [
    "Руководитель ФИО (DaData)",
    "Должность руководителя (DaData)",
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
]

# ===================== Утилиты =====================
def normalize_str(s):
    return str(s).strip() if s is not None else ""

def normalize_site(site):
    if not site:
        return ""
    s = str(site).strip().lower()
    for p in ["https://", "http://", "www."]:
        if s.startswith(p):
            s = s[len(p):]
    return s.rstrip("/").strip()

def normalize_email(email):
    return str(email).strip().lower() if email else ""

def extract_fields(dadata_resp):
    """Вытаскивает нужное из ответа DaData"""
    if not dadata_resp or dadata_resp.get("http_status") != 200 or not dadata_resp.get("data"):
        return {
            "fio": "", "post": "", "email": "", "site": "", "phone": "",
            "ogrn": "", "full_name": "", "status": "", "actuality": "",
            "raw_status": dadata_resp.get("error", "Нет данных")[:80] if dadata_resp else "Нет данных"
        }
    sugg = dadata_resp["data"].get("suggestions", [])
    if not sugg:
        return {
            "fio": "", "post": "", "email": "", "site": "", "phone": "",
            "ogrn": "", "full_name": "", "status": "Не найдено", "actuality": "",
            "raw_status": "Не найдено"
        }
    d = sugg[0].get("data", {}) or {}
    mgmt = d.get("management") or {}
    state = d.get("state") or {}
    actuality = state.get("actuality_date") or state.get("registration_date") or ""
    # форматирование даты из timestamp
    if isinstance(actuality, int) and actuality > 1000000000:
        try:
            actuality = datetime.fromtimestamp(actuality/1000).strftime("%Y-%m-%d")
        except:
            pass
    return {
        "fio": normalize_str(mgmt.get("name")),
        "post": normalize_str(mgmt.get("post")),
        "email": normalize_str(d.get("email")),
        "site": normalize_str(d.get("site")),
        "phone": normalize_str(d.get("phone")),
        "ogrn": normalize_str(d.get("ogrn")),
        "full_name": normalize_str(sugg[0].get("unrestricted_value") or sugg[0].get("value", "")),
        "status": normalize_str(state.get("status")),
        "actuality": normalize_str(actuality),
        "raw_status": "Найдено",
    }

def call_dadata(inn, token, session, retries=3):
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Authorization": f"Token {token}",
    }
    last_err = None
    for attempt in range(retries):
        try:
            resp = session.post(DADATA_URL, json={"query": inn}, headers=headers, timeout=15)
            if resp.status_code == 429:
                wait = 2 ** attempt + 1
                time.sleep(wait)
                continue
            if resp.status_code != 200:
                return {"inn": inn, "http_status": resp.status_code, "error": resp.text[:2000]}
            return {"inn": inn, "http_status": 200, "data": resp.json()}
        except requests.exceptions.RequestException as e:
            last_err = str(e)
            time.sleep(1.2 ** attempt)
    return {"inn": inn, "http_status": None, "error": last_err or "max retries"}

# ===================== GUI =====================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DaData — Парсер Администраций городов (Windows GUI)")
        self.geometry("980x720")
        self.minsize(940, 680)
        try:
            self.iconbitmap(default="")
        except:
            pass

        self.token_var = tk.StringVar(value=DEFAULT_TOKEN)
        self.input_var = tk.StringVar(value=DEFAULT_INPUT)
        self.output_var = tk.StringVar(value=DEFAULT_OUTPUT)
        self.raw_var = tk.StringVar(value=DEFAULT_RAW)
        self.delay_var = tk.DoubleVar(value=0.2)
        self.resume_var = tk.BooleanVar(value=True)
        self.status_text = tk.StringVar(value="Готов к запуску")

        self.is_running = False
        self.should_stop = False
        self.worker_thread = None
        self.log_queue = queue.Queue()
        self.results = {}  # inn -> response
        self.stats = {"total": 0, "done": 0, "found": 0, "errors": 0}

        self.create_widgets()
        self.after(100, self.process_log_queue)
        # Проверка наличия входного файла
        self.check_input_file()

    def create_widgets(self):
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except:
            style.theme_use("clam")
        style.configure("TButton", padding=6)
        style.configure("Header.TLabel", font=("Segoe UI", 14, "bold"), foreground="#1F4E78")
        style.configure("Sub.TLabel", font=("Segoe UI", 8), foreground="#666")
        style.configure("Card.TFrame", background="white")

        # Верхний заголовок
        header = ttk.Frame(self, padding=(16, 12))
        header.pack(fill="x")
        ttk.Label(header, text="🔍 DaData Find-Party — Администрации городов", style="Header.TLabel").pack(anchor="w")
        ttk.Label(header, text="Находит по ИНН руководителя + должность, проверяет email/сайт и сохраняет каждый сырой ответ (JSONL) для дозабора любых полей", style="Sub.TLabel", wraplength=940).pack(anchor="w", pady=(4,0))

        # Карточка настроек
        card = ttk.LabelFrame(self, text="  Настройки  ", padding=12)
        card.pack(fill="x", padx=16, pady=(8, 6))

        # Строка 1: входной файл
        row1 = ttk.Frame(card)
        row1.pack(fill="x", pady=4)
        ttk.Label(row1, text="Входной Excel:", width=16).pack(side="left")
        ttk.Entry(row1, textvariable=self.input_var).pack(side="left", fill="x", expand=True, padx=(6, 6))
        ttk.Button(row1, text="Обзор…", command=self.browse_input, width=10).pack(side="left")
        ttk.Button(row1, text="Открыть папку", command=self.open_folder).pack(side="left", padx=(6,0))

        # Строка 2: выходной и raw
        row2 = ttk.Frame(card)
        row2.pack(fill="x", pady=4)
        ttk.Label(row2, text="Выходной Excel:", width=16).pack(side="left")
        ttk.Entry(row2, textvariable=self.output_var).pack(side="left", fill="x", expand=True, padx=(6, 6))
        ttk.Button(row2, text="Обзор…", command=self.browse_output, width=10).pack(side="left")

        row3 = ttk.Frame(card)
        row3.pack(fill="x", pady=4)
        ttk.Label(row3, text="Сырые ответы:", width=16).pack(side="left")
        ttk.Entry(row3, textvariable=self.raw_var).pack(side="left", fill="x", expand=True, padx=(6, 6))
        ttk.Button(row3, text="Обзор…", command=self.browse_raw, width=10).pack(side="left")
        ttk.Label(row3, text="JSONL", foreground="#888").pack(side="left", padx=(6,0))

        # Строка 4: токен и задержка
        row4 = ttk.Frame(card)
        row4.pack(fill="x", pady=4)
        ttk.Label(row4, text="DaData токен:", width=16).pack(side="left")
        self.token_entry = ttk.Entry(row4, textvariable=self.token_var, show="•", width=40)
        self.token_entry.pack(side="left", fill="x", expand=True, padx=(6, 6))
        self.show_token_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row4, text="Показать", variable=self.show_token_var, command=self.toggle_token).pack(side="left")
        ttk.Button(row4, text="Проверить", command=self.test_token, width=10).pack(side="left", padx=(6,0))

        row5 = ttk.Frame(card)
        row5.pack(fill="x", pady=4)
        ttk.Label(row5, text="Пауза, сек:", width=16).pack(side="left")
        ttk.Spinbox(row5, from_=0.05, to=1.0, increment=0.05, textvariable=self.delay_var, width=8, format="%.2f").pack(side="left", padx=(6, 12))
        ttk.Label(row5, text="0.15–0.20 = 5–6 запр/сек, защита от 429", foreground="#888").pack(side="left")
        ttk.Checkbutton(row5, text="Resume (продолжить с места остановки)", variable=self.resume_var).pack(side="left", padx=(16,0))

        row6 = ttk.Frame(card)
        row6.pack(fill="x", pady=4)
        ttk.Label(row6, text="API:", width=16).pack(side="left")
        link = ttk.Label(row6, text="https://dadata.ru/api/find-party/  (POST suggestions.dadata.ru/suggestions/api/4_1/rs/findById/party)", foreground="#1F4E78", cursor="hand2")
        link.pack(side="left", padx=(6,0))
        link.bind("<Button-1>", lambda e: webbrowser.open("https://dadata.ru/api/find-party/"))

        # Статистика + прогресс
        stats = ttk.Frame(self, padding=(16, 6))
        stats.pack(fill="x")
        self.stat_total = ttk.Label(stats, text="Всего: 0", font=("Segoe UI", 9, "bold"))
        self.stat_total.pack(side="left", padx=(0, 16))
        self.stat_done = ttk.Label(stats, text="Готово: 0")
        self.stat_done.pack(side="left", padx=(0, 16))
        self.stat_found = ttk.Label(stats, text="Найдено: 0", foreground="#15803d")
        self.stat_found.pack(side="left", padx=(0, 16))
        self.stat_err = ttk.Label(stats, text="Ошибок: 0", foreground="#b91c1c")
        self.stat_err.pack(side="left")
        self.stat_status = ttk.Label(stats, textvariable=self.status_text, foreground="#1F4E78")
        self.stat_status.pack(side="right")

        self.progress = ttk.Progressbar(self, mode="determinate")
        self.progress.pack(fill="x", padx=16, pady=(4, 8))

        # Кнопки управления
        btns = ttk.Frame(self, padding=(16, 0))
        btns.pack(fill="x")
        self.start_btn = ttk.Button(btns, text="▶  Запустить обогащение", command=self.start, style="TButton")
        self.start_btn.pack(side="left", padx=(0, 8))
        self.stop_btn = ttk.Button(btns, text="⏸  Стоп", command=self.stop, state="disabled")
        self.stop_btn.pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="🧪 Тест 1 запроса (Пермь 5902290635)", command=self.test_one).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="📂 Открыть результат", command=self.open_result).pack(side="left")
        ttk.Label(btns, text="Лог сохраняет каждый ответ — можно дозабрать любые поля без повторных запросов", foreground="#888").pack(side="left", padx=(12,0))

        # Лог + таблица
        paned = ttk.PanedWindow(self, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=16, pady=8)

        # Лог
        log_frame = ttk.LabelFrame(paned, text="  Лог  ", padding=6)
        self.log_text = tk.Text(log_frame, height=18, font=("Consolas", 9), wrap="word", bg="#0f172a", fg="#e2e8f0", insertbackground="white")
        self.log_text.pack(side="left", fill="both", expand=True)
        log_scroll = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        log_scroll.pack(side="right", fill="y")
        self.log_text.config(yscrollcommand=log_scroll.set)
        paned.add(log_frame, weight=3)

        # Таблица предпросмотра
        table_frame = ttk.LabelFrame(paned, text="  Предпросмотр (последние 50)  ", padding=6)
        cols = ("inn", "fio", "post", "status")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", height=18)
        self.tree.heading("inn", text="ИНН")
        self.tree.heading("fio", text="Руководитель")
        self.tree.heading("post", text="Должность")
        self.tree.heading("status", text="Статус")
        self.tree.column("inn", width=110, anchor="center")
        self.tree.column("fio", width=170)
        self.tree.column("post", width=170)
        self.tree.column("status", width=90, anchor="center")
        self.tree.pack(side="left", fill="both", expand=True)
        tree_scroll = ttk.Scrollbar(table_frame, command=self.tree.yview)
        tree_scroll.pack(side="right", fill="y")
        self.tree.config(yscrollcommand=tree_scroll.set)
        paned.add(table_frame, weight=2)

        # Низ
        footer = ttk.Frame(self, padding=(16, 6))
        footer.pack(fill="x")
        ttk.Label(footer, text="Подсказка: сырые ответы в JSONL — из них потом вытяните адрес, ОКВЭД, учредителей и т.д. без новых запросов к DaData", foreground="#888", font=("Segoe UI", 8)).pack(side="left")
        ttk.Label(footer, text="ветка arena/363bbf0c-gov", foreground="#aaa", font=("Segoe UI", 8)).pack(side="right")

        self.log("Готов. Проверьте пути и токен, затем нажмите «Запустить».")
        self.log(f"Входной по умолчанию: {DEFAULT_INPUT}")
        self.log("Токен предзаполнен вашим ключом 77eff...f382dfb — можете заменить на свой.")

    # ===== Helpers =====
    def log(self, msg, level="info"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_queue.put(f"[{timestamp}] {msg}")

    def process_log_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_text.insert("end", msg + "\n")
                self.log_text.see("end")
        except queue.Empty:
            pass
        self.after(100, self.process_log_queue)

    def check_input_file(self):
        p = Path(self.input_var.get())
        if not p.exists():
            self.log(f"⚠ Входной файл не найден: {p} — выберите через Обзор", level="warn")
            self.status_text.set(f"Не найден: {p.name}")
        else:
            try:
                wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
                sheets = [ws.title for ws in wb.worksheets if ws.title != "Оглавление"]
                total_rows = sum(ws.max_row - 1 for ws in wb.worksheets if ws.title != "Оглавление")
                self.log(f"Найден входной файл: {p} — листов: {sheets}")
                self.status_text.set(f"Готов: {total_rows} строк в {len(sheets)} листах")
            except Exception as e:
                self.log(f"Ошибка чтения Excel: {e}")

    def browse_input(self):
        path = filedialog.askopenfilename(title="Выберите входной Excel", filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")])
        if path:
            self.input_var.set(path)
            self.check_input_file()

    def browse_output(self):
        path = filedialog.asksaveasfilename(title="Сохранить как", defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx"), ("All", "*.*")], initialfile=Path(self.output_var.get()).name)
        if path:
            self.output_var.set(path)

    def browse_raw(self):
        path = filedialog.asksaveasfilename(title="Файл сырых ответов", defaultextension=".jsonl", filetypes=[("JSONL", "*.jsonl"), ("All", "*.*")], initialfile=Path(self.raw_var.get()).name)
        if path:
            self.raw_var.set(path)

    def open_folder(self):
        p = Path(self.input_var.get()).parent if Path(self.input_var.get()).parent.exists() else Path.cwd()
        try:
            os.startfile(p)
        except:
            webbrowser.open(str(p))

    def open_result(self):
        p = Path(self.output_var.get())
        if p.exists():
            try:
                os.startfile(p)
            except:
                webbrowser.open(str(p.parent))
        else:
            messagebox.showinfo("Открыть результат", f"Файл ещё не создан:\n{p}\nПосле завершения он появится здесь.")

    def toggle_token(self):
        self.token_entry.config(show="" if self.show_token_var.get() else "•")

    def test_token(self):
        token = self.token_var.get().strip()
        if not token:
            messagebox.showwarning("Токен", "Введите токен DaData")
            return
        self.log(f"Проверка токена {token[:6]}...{token[-4:]} — запрос к DaData (ИНН 7707083893)...")
        def do_test():
            try:
                r = requests.post(DADATA_URL, json={"query": "7707083893"}, headers={"Content-Type": "application/json", "Accept": "application/json", "Authorization": f"Token {token}"}, timeout=10)
                if r.status_code == 200 and r.json().get("suggestions"):
                    self.log("✔ Токен рабочий — DaData ответила (найдено)")
                    self.status_text.set("Токен OK")
                elif r.status_code == 200:
                    self.log("✔ Токен рабочий, но организация не найдена (пустой ответ)")
                elif r.status_code == 401:
                    self.log("✘ 401 — неверный токен")
                elif r.status_code == 403:
                    self.log("✘ 403 — нет доступа, проверьте тариф")
                else:
                    self.log(f"✘ HTTP {r.status_code}: {r.text[:300]}")
            except Exception as e:
                self.log(f"✘ Ошибка проверки: {e}")
        threading.Thread(target=do_test, daemon=True).start()

    def test_one(self):
        token = self.token_var.get().strip()
        if not token:
            messagebox.showwarning("Токен", "Введите токен")
            return
        self.log("Тест: ИНН 5902290635 (Администрация Перми)...")
        def do():
            try:
                resp = call_dadata("5902290635", token, requests.Session())
                if resp.get("http_status") == 200 and resp.get("data", {}).get("suggestions"):
                    d = resp["data"]["suggestions"][0]["data"]
                    fio = d.get("management", {}).get("name", "")
                    post = d.get("management", {}).get("post", "")
                    self.log(f"✔ Найдено: {fio} — {post} (status={d.get('state', {}).get('status')})")
                    self.tree.insert("", "end", values=("5902290635", fio, post, "тест OK"))
                else:
                    self.log(f"Ответ: {resp}")
            except Exception as e:
                self.log(f"✘ {e}")
        threading.Thread(target=do, daemon=True).start()

    # ===== Main logic =====
    def start(self):
        if self.is_running:
            return
        input_path = Path(self.input_var.get())
        if not input_path.exists():
            messagebox.showerror("Файл не найден", f"Не найден входной файл:\n{input_path}")
            return
        token = self.token_var.get().strip()
        if not token:
            messagebox.showerror("Токен", "Введите API ключ DaData (https://dadata.ru/api/find-party/)")
            return
        self.is_running = True
        self.should_stop = False
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.log_text.delete("1.0", "end")
        self.log(f"Старт. Вход: {input_path}")
        self.log(f"Выход: {self.output_var.get()} | Сырые: {self.raw_var.get()} | Пауза: {self.delay_var.get()}с | Resume: {self.resume_var.get()}")
        self.worker_thread = threading.Thread(target=self.worker, daemon=True)
        self.worker_thread.start()

    def stop(self):
        if not self.is_running:
            return
        self.should_stop = True
        self.log("⏸ Остановка после текущего запроса...")
        self.status_text.set("Останавливается...")

    def worker(self):
        input_path = Path(self.input_var.get())
        output_path = Path(self.output_var.get())
        raw_path = Path(self.raw_var.get())
        token = self.token_var.get().strip()
        delay = float(self.delay_var.get())
        resume = bool(self.resume_var.get())

        # Загрузка ИНН
        try:
            wb = openpyxl.load_workbook(input_path, read_only=True, data_only=True)
        except Exception as e:
            self.log(f"✘ Не удалось открыть Excel: {e}")
            self.finish(False)
            return

        inn_to_rows = defaultdict(list)
        inn_set = set()
        inn_examples = {}  # inn -> example full_name
        sheets = [ws for ws in wb.worksheets if ws.title != "Оглавление"]
        for ws in sheets:
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue
            header = rows[0]
            try:
                inn_idx = header.index("ИНН")
            except:
                inn_idx = 5
            for r_idx, row in enumerate(rows[1:], start=2):
                inn = normalize_str(row[inn_idx])
                if inn and inn.isdigit() and len(inn) in (10, 12):
                    inn_set.add(inn)
                    inn_to_rows[inn].append((ws.title, r_idx))
                    if inn not in inn_examples:
                        inn_examples[inn] = normalize_str(row[0])[:80]

        wb.close()
        total_unique = len(inn_set)
        self.log(f"Найдено уникальных ИНН: {total_unique} (из {sum(len(list(s.iter_rows()))-1 for s in sheets) if False else '—'} строк)")
        self.stats["total"] = total_unique
        self.update_stats()
        self.progress.config(maximum=total_unique)

        # Resume: загрузить уже сохранённые
        existing = {}
        if resume and raw_path.exists():
            self.log(f"Resume: найден {raw_path} — подгружаю...")
            try:
                with open(raw_path, "r", encoding="utf-8") as f:
                    for line in f:
                        try:
                            obj = json.loads(line)
                            inn = obj.get("inn")
                            if inn:
                                existing[inn] = obj.get("response") or obj
                        except:
                            continue
                self.log(f"  Загружено {len(existing)} уже запрошенных ИНН — пропускаем их")
                # восстановить статистику для уже готовых
                for inn in existing:
                    if inn in inn_set:
                        self.stats["done"] += 1
                        # проверить found
                        resp = existing[inn]
                        if isinstance(resp, dict) and resp.get("data", {}).get("suggestions"):
                            self.stats["found"] += 1
                self.update_stats()
            except Exception as e:
                self.log(f"  Ошибка чтения resume: {e}")

        to_query = [inn for inn in sorted(inn_set) if inn not in existing]
        self.log(f"К запросу осталось: {len(to_query)} ИНН")
        if not to_query:
            self.log("Все ИНН уже запрошены — сразу собираю Excel")
            # Загрузить из уже сохранённого raw
            full_map = {}
            if raw_path.exists():
                with open(raw_path, "r", encoding="utf-8") as f:
                    for line in f:
                        try:
                            obj = json.loads(line)
                            full_map[obj["inn"]] = obj["response"]
                        except:
                            continue
            self.results = full_map
        else:
            session = requests.Session()
            session.headers.update({"User-Agent": "DaData-GUI/1.0 (Windows)"})

            # Открываем raw для дозаписи
            raw_file = open(raw_path, "a", encoding="utf-8") if resume else open(raw_path, "w", encoding="utf-8")
            # Сохраним уже существующие в файл, если файл был пустой? Нет, они уже там.

            start_time = time.time()
            for idx, inn in enumerate(to_query, start=1):
                if self.should_stop:
                    self.log("Остановлено пользователем")
                    raw_file.close()
                    self.finish(False)
                    return

                # Прогресс
                overall_done = self.stats["done"] + 1
                self.log(f"[{idx}/{len(to_query)}] {inn} ({inn_examples.get(inn, '')[:50]}...)")
                resp = call_dadata(inn, token, session)
                # Сохранить
                record = {"inn": inn, "queried_at": datetime.now().isoformat(), "response": resp}
                # Для raw используем формат {inn, queried_at, response}
                raw_file.write(json.dumps(record, ensure_ascii=False) + "\n")
                raw_file.flush()

                # Обновить статистику
                self.stats["done"] += 1
                if resp.get("http_status") == 200 and resp.get("data", {}).get("suggestions"):
                    self.stats["found"] += 1
                    # Добавить в таблицу
                    d = resp["data"]["suggestions"][0].get("data", {})
                    mgmt = d.get("management", {}) or {}
                    fio = mgmt.get("name", "")
                    post = mgmt.get("post", "")
                    self.add_tree_row(inn, fio, post, "найдено")
                    if idx % 10 == 0:
                        self.log(f"  → {fio} — {post}")
                elif resp.get("http_status") == 200:
                    self.add_tree_row(inn, "—", "—", "не найдено")
                else:
                    self.stats["errors"] += 1
                    err = resp.get("error", str(resp.get("http_status")))
                    self.log(f"  ✘ Ошибка: {err[:120]}")
                    self.add_tree_row(inn, "ошибка", err[:30], f"HTTP {resp.get('http_status')}")

                self.update_stats()
                elapsed = time.time() - start_time
                rps = idx / elapsed if elapsed > 0 else 0
                self.status_text.set(f"{self.stats['done']}/{total_unique} • {rps:.1f} req/s • {elapsed/60:.1f} мин")
                self.progress["value"] = self.stats["done"]

                if idx % 50 == 0:
                    self.log(f"Прогресс {idx}/{len(to_query)} — {rps:.1f} r/s")

                time.sleep(delay)

                if idx % 200 == 0:
                    raw_file.flush()
                    try:
                        os.fsync(raw_file.fileno())
                    except:
                        pass

            raw_file.close()
            # Перечитаем raw полностью для Excel
            full_map = {}
            with open(raw_path, "r", encoding="utf-8") as f:
                for line in f:
                    try:
                        obj = json.loads(line)
                        full_map[obj["inn"]] = obj["response"]
                    except:
                        continue
            self.results = full_map

        # Генерация Excel
        self.log("Собираю обогащённый Excel...")
        self.status_text.set("Генерация Excel...")
        try:
            out = self.generate_excel(input_path, output_path, self.results)
            self.log(f"✔ Excel сохранён: {out} ({Path(out).stat().st_size/1024:.1f} КБ)")
            self.log(f"✔ Сырые ответы: {raw_path} ({Path(raw_path).stat().st_size/1024:.1f} КБ)")
            self.log("Готово! Откройте Excel — новые колонки зелёные справа, фильтры включены.")
            self.status_text.set("Готово")
            self.progress["value"] = total_unique
        except Exception as e:
            self.log(f"✘ Ошибка генерации Excel: {e}")
            import traceback
            self.log(traceback.format_exc())
            self.status_text.set("Ошибка Excel")

        self.finish(True)

    def add_tree_row(self, inn, fio, post, status):
        # Вызывается из worker thread — нужно через queue или after
        def do():
            # Ограничить до 50 видимых
            if len(self.tree.get_children()) >= 50:
                self.tree.delete(self.tree.get_children()[0])
            self.tree.insert("", "end", values=(inn, fio, post, status))
            self.tree.yview_moveto(1.0)
        self.after(0, do)

    def update_stats(self):
        def do():
            self.stat_total.config(text=f"Всего: {self.stats['total']}")
            self.stat_done.config(text=f"Готово: {self.stats['done']}")
            self.stat_found.config(text=f"Найдено: {self.stats['found']}")
            self.stat_err.config(text=f"Ошибок: {self.stats['errors']}")
            self.progress["value"] = self.stats["done"]
        self.after(0, do)

    def generate_excel(self, input_path, output_path, results):
        """Генерация обогащённого Excel (копия логики из enrich_dadata.py)"""
        wb = openpyxl.load_workbook(input_path, data_only=True)

        HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        NEW_FILL = PatternFill(start_color="548235", end_color="548235", fill_type="solid")
        HEADER_FONT = Font(color="FFFFFF", bold=True, size=10, name="Calibri")
        THIN_BORDER = Border(left=Side(style='thin', color='B4C6E7'), right=Side(style='thin', color='B4C6E7'), top=Side(style='thin', color='B4C6E7'), bottom=Side(style='thin', color='B4C6E7'))
        ALT_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

        for ws in wb.worksheets:
            if ws.title == "Оглавление":
                continue
            header_len = ws.max_column
            # Добавить заголовки
            for idx, name in enumerate(NEW_COLUMNS, start=header_len+1):
                cell = ws.cell(row=1, column=idx, value=name)
                cell.fill = NEW_FILL
                cell.font = HEADER_FONT
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.border = THIN_BORDER
            for cell in ws[1]:
                if cell.column <= header_len:
                    cell.fill = HEADER_FILL
                    cell.font = HEADER_FONT
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    cell.border = THIN_BORDER
            ws.row_dimensions[1].height = 38
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
            ws.page_setup.paperSize = ws.PAPERSIZE_A3
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            widths = [55,30,30,32,16,14,13,32,18,26,22,32,16,18, 26,28,26,26,18, 18,18,18,18, 16,26,28,16,16]
            for i, w in enumerate(widths, start=1):
                if i <= ws.max_column:
                    ws.column_dimensions[get_column_letter(i)].width = w

            for r_idx in range(2, ws.max_row+1):
                inn = normalize_str(ws.cell(row=r_idx, column=6).value)
                email_reg = normalize_str(ws.cell(row=r_idx, column=10).value)
                site_reg = normalize_str(ws.cell(row=r_idx, column=8).value)
                resp = results.get(inn, {})
                fields = extract_fields(resp)
                fio = fields["fio"]
                post = fields["post"]
                email_dd = fields["email"]
                site_dd = fields["site"]
                phone_dd = fields["phone"]
                ogrn_dd = fields["ogrn"]
                full_dd = fields["full_name"]
                status_dd = fields["status"]
                actuality = fields["actuality"]
                raw_status = fields["raw_status"]

                email_match = "Да" if email_dd and normalize_email(email_dd) == normalize_email(email_reg) else ("Нет" if email_dd and email_reg else ("Пусто в DaData" if not email_dd else "Пусто в реестре"))
                site_match = "Да" if site_dd and normalize_site(site_dd) == normalize_site(site_reg) else ("Нет" if site_dd and site_reg else ("Пусто в DaData" if not site_dd else "Пусто в реестре"))
                has_extra_email = "Да" if email_dd and normalize_email(email_dd) != normalize_email(email_reg) and email_dd else "Нет"
                has_extra_site = "Да" if site_dd and normalize_site(site_dd) != normalize_site(site_reg) and site_dd else "Нет"

                vals = [fio, post, email_dd, site_dd, phone_dd, email_match, site_match, has_extra_email, has_extra_site, ogrn_dd, full_dd, status_dd, actuality, raw_status, datetime.now().strftime("%Y-%m-%d %H:%M")]
                for off, val in enumerate(vals, start=0):
                    ws.cell(row=r_idx, column=header_len+1+off, value=val)

            for r_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
                is_alt = (r_idx % 2 == 0)
                for cell in row:
                    cell.alignment = Alignment(vertical='center', wrap_text=True)
                    cell.border = THIN_BORDER
                    cell.font = Font(size=9, name="Calibri")
                    if is_alt:
                        cell.fill = ALT_FILL
                ws.row_dimensions[r_idx].height = 16

        wb.save(output_path)
        return output_path

    def finish(self, success):
        def do():
            self.is_running = False
            self.start_btn.config(state="normal")
            self.stop_btn.config(state="disabled")
            if success:
                self.start_btn.config(text="✅ Готово — запустить снова")
                messagebox.showinfo("Готово", f"Обогащение завершено!\n\nExcel: {self.output_var.get()}\nСырые: {self.raw_var.get()}\n\nНовые колонки — зелёные справа.\nСырые JSONL — для дозабора любых полей.")
            else:
                self.start_btn.config(text="▶  Запустить снова")
        self.after(0, do)

if __name__ == "__main__":
    # Windows: поддержка кириллицы в консоли
    if sys.platform == "win32":
        try:
            os.system("chcp 65001 > nul")
        except:
            pass
    app = App()
    app.mainloop()
