"""
GUI для Dadata Parser — простой запуск без команд.
Вся настройка (токены, прокси, файлы) — внутри окна.
Первый запуск: автопроверка зависимостей с кнопкой «Установить всё».
"""
from __future__ import annotations
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pathlib, threading, queue, sys, subprocess, importlib.util
from datetime import datetime

# --- dependency check ---
REQUIRED_PKGS = ["openpyxl", "requests", "yaml", "tqdm"]
def missing_deps():
    miss=[]
    for m in ["openpyxl", "requests"]:
        if importlib.util.find_spec(m) is None:
            miss.append(m)
    if importlib.util.find_spec("yaml") is None:
        miss.append("pyyaml")
    if importlib.util.find_spec("socks") is None:
        miss.append("requests[socks]")
    if importlib.util.find_spec("tqdm") is None:
        miss.append("tqdm")
    return miss

def load_cfg():
    from ..config import load_config
    return load_config()

def save_cfg(cfg, path=None):
    from ..config import save_config
    save_config(cfg, path)

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DaData Parser — простой запуск")
        self.geometry("1120x820")
        self.minsize(1020, 680)
        try: self.iconbitmap(default="")
        except: pass

        self.cfg=load_cfg()
        self.log_queue=queue.Queue()
        self.is_running=False
        self.create_widgets()
        self.after(100, self.poll_log)
        self.check_deps_banner()
        self.refresh_token_stats()
        # приветствие первого запуска
        tokens=self.cfg.get("dadata",{}).get("tokens",[])
        if not tokens or tokens==[""]:
            self.after(500, self.show_welcome)

    def create_widgets(self):
        style=ttk.Style()
        try: style.theme_use("vista")
        except: style.theme_use("clam")
        style.configure("Header.TLabel", font=("Segoe UI", 15, "bold"), foreground="#0B5394")
        style.configure("Step.TLabel", font=("Segoe UI", 9, "bold"), foreground="#ffffff", background="#0B5394")
        style.configure("Card.TLabelframe.Label", font=("Segoe UI", 9, "bold"))
        style.configure("Big.TButton", font=("Segoe UI", 11, "bold"))

        # Шапка
        header=ttk.Frame(self, padding=(16,10))
        header.pack(fill="x")
        ttk.Label(header, text="🔍 DaData Parser", style="Header.TLabel").pack(side="left")
        ttk.Label(header, text="  —  токены и прокси настраиваются прямо здесь, cmd не нужен", font=("Segoe UI", 9), foreground="#666").pack(side="left", padx=(8,0))
        ttk.Label(header, text="API: suggestions.dadata.ru/findById/party → 10 000/сутки/токен → авторотация", font=("Consolas", 7), foreground="#888").pack(side="right")

        # Баннер зависимостей (показывается если чего-то нет)
        self.deps_banner=ttk.Frame(self, padding=(12,6))
        self.deps_label=ttk.Label(self.deps_banner, text="", foreground="#b91c1c")
        self.deps_label.pack(side="left")
        self.deps_btn=ttk.Button(self.deps_banner, text="📦 Установить всё (1 клик)", command=self.install_deps)
        self.deps_btn.pack(side="right", padx=(8,0))

        # Шаги-подсказка
        steps=ttk.Frame(self, padding=(12,0))
        steps.pack(fill="x")
        for i, txt in enumerate(["Шаг 1: Вставьте токены DaData", "Шаг 2: Вставьте прокси (или выключите)", "Шаг 3: Выберите файл и Запустить"],1):
            f=tk.Frame(steps, bg="#0B5394", padx=10, pady=3)
            f.pack(side="left", padx=(0,8))
            tk.Label(f, text=f"{i}", font=("Segoe UI", 9, "bold"), bg="#ffffff", fg="#0B5394", width=2).pack(side="left", padx=(0,6))
            tk.Label(f, text=txt, font=("Segoe UI", 8), bg="#0B5394", fg="white").pack(side="left")

        paned=ttk.PanedWindow(self, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=12, pady=8)

        # Левая панель
        left=ttk.LabelFrame(paned, text=" ⚙️ Настройки — всё здесь ", padding=10)
        paned.add(left, weight=1)

        # Токены — большой блок
        ttl=ttk.Frame(left)
        ttl.pack(fill="x")
        ttk.Label(ttl, text="Токены DaData", font=("Segoe UI", 9, "bold")).pack(side="left")
        ttk.Label(ttl, text="каждый до 10 000/сутки — исчерпан → авто следующий", font=("Segoe UI", 7), foreground="#888").pack(side="left", padx=(8,0))
        ttk.Button(ttl, text="где взять?", width=10, command=lambda: self.log("Токен: https://dadata.ru/profile/#info — скопируйте API-ключ")).pack(side="right")

        self.tokens_text=scrolledtext.ScrolledText(left, height=5, font=("Consolas", 9), wrap="none")
        self.tokens_text.pack(fill="x", pady=(4,2))
        self.tokens_text.configure(bg="#fffff0")
        for t in self.cfg.get("dadata",{}).get("tokens",[]):
            self.tokens_text.insert("end", t+"\n")
        # placeholder если пусто
        if not self.tokens_text.get("1.0","end").strip():
            self.tokens_text.insert("1.0", "77eff8edd3640fc0002ad6ae2431261a2f382dfb\n# второй токен можно с новой строки\n")
            self.tokens_text.configure(fg="#888")
            def clear_ph(e):
                if "второй токен" in self.tokens_text.get("1.0","end"):
                    self.tokens_text.delete("1.0","end")
                    self.tokens_text.configure(fg="black")
            self.tokens_text.bind("<FocusIn>", clear_ph)

        ttk.Label(left, text="💡 Вставьте каждый токен с новой строки. Лимит считается по МСК, следующий слот в 00:00.", font=("Segoe UI", 7), foreground="#666").pack(anchor="w")

        row_lim=ttk.Frame(left)
        row_lim.pack(fill="x", pady=(6,0))
        ttk.Label(row_lim, text="Лимит/сутки/токен:").pack(side="left")
        self.limit_var=tk.IntVar(value=self.cfg["dadata"].get("daily_limit_per_token",10000))
        ttk.Spinbox(row_lim, from_=100, to=50000, increment=1000, textvariable=self.limit_var, width=9).pack(side="left", padx=(6,0))
        ttk.Label(row_lim, text="10 000 — по тарифу DaData", foreground="#888", font=("Segoe UI",7)).pack(side="left", padx=(6,0))

        # Прокси
        ttk.Separator(left).pack(fill="x", pady=8)
        row_p=ttk.Frame(left)
        row_p.pack(fill="x")
        self.proxy_enabled=tk.BooleanVar(value=self.cfg.get("proxy",{}).get("enabled",True))
        ttk.Checkbutton(row_p, text="Использовать SOCKS5 прокси", variable=self.proxy_enabled, command=self.toggle_proxy).pack(side="left")
        ttk.Label(row_p, text="формат IP:PORT:LOGIN:PASS", font=("Consolas",7), foreground="#888").pack(side="left", padx=(8,0))
        self.proxy_text=scrolledtext.ScrolledText(left, height=3, font=("Consolas", 9))
        self.proxy_text.pack(fill="x", pady=(4,2))
        for p in self.cfg.get("proxy",{}).get("list",[]):
            self.proxy_text.insert("end", p+"\n")
        if not self.proxy_text.get("1.0","end").strip():
            self.proxy_text.insert("1.0", "95.78.126.13:14755:5nqb8nr:9cz93bt\n")
            self.proxy_text.configure(fg="#888")
            def clear_pp(e):
                if "95.78.126.13" in self.proxy_text.get("1.0","end"):
                    if self.proxy_text.get("1.0","end").strip().startswith("95.78"):
                        self.proxy_text.delete("1.0","end")
                        self.proxy_text.configure(fg="black")
            self.proxy_text.bind("<FocusIn>", clear_pp)
        ttk.Label(left, text="Пул может быть 1 прокси с авторотацией — провайдер сам меняет IP.", font=("Segoe UI", 7), foreground="#666").pack(anchor="w")
        self.proxy_hint=ttk.Label(left, text="", font=("Segoe UI",7), foreground="#0B5394")
        self.proxy_hint.pack(anchor="w")

        # Файлы
        ttk.Separator(left).pack(fill="x", pady=8)
        ttk.Label(left, text="Файлы", font=("Segoe UI", 9, "bold")).pack(anchor="w")
        # вход
        row_in=ttk.Frame(left)
        row_in.pack(fill="x", pady=(4,0))
        ttk.Label(row_in, text="Откуда ИНН:", width=12).pack(side="left")
        self.input_var=tk.StringVar(value=self.cfg.get("output",{}).get("excel_input") or "")
        self.input_entry=ttk.Entry(row_in, textvariable=self.input_var)
        self.input_entry.pack(side="left", fill="x", expand=True, padx=(6,6))
        ttk.Button(row_in, text="📂 Выбрать", width=10, command=self.browse_input).pack(side="left")
        ttk.Label(left, text="Подходит: .xlsx (колонка ИНН) / .txt / .csv / .jsonl — можно перетащить файл в окно", font=("Segoe UI",7), foreground="#888").pack(anchor="w", pady=(2,0))

        row_out=ttk.Frame(left)
        row_out.pack(fill="x", pady=6)
        ttk.Label(row_out, text="Куда сохранить:", width=12).pack(side="left")
        self.output_var=tk.StringVar(value=self.cfg.get("output",{}).get("excel_output") or "./data/enriched.xlsx")
        ttk.Entry(row_out, textvariable=self.output_var).pack(side="left", fill="x", expand=True, padx=(6,6))
        ttk.Button(row_out, text="📂 Выбрать", width=10, command=self.browse_output).pack(side="left")

        row_delay=ttk.Frame(left)
        row_delay.pack(fill="x", pady=4)
        ttk.Label(row_delay, text="Пауза:").pack(side="left")
        self.delay_var=tk.DoubleVar(value=self.cfg["dadata"].get("delay_between_requests",0.2))
        ttk.Spinbox(row_delay, from_=0.05, to=2.0, increment=0.05, textvariable=self.delay_var, width=7, format="%.2f").pack(side="left", padx=(6,0))
        ttk.Label(row_delay, text="сек (0.2 = 5 r/s, защита от 429)", font=("Segoe UI",7), foreground="#888").pack(side="left", padx=(6,0))

        # Действия
        row_btn=ttk.Frame(left)
        row_btn.pack(fill="x", pady=8)
        ttk.Button(row_btn, text="💾 Сохранить", command=self.save_config).pack(side="left", padx=(0,6))
        ttk.Button(row_btn, text="🧪 Тест прокси", command=self.test_proxy).pack(side="left", padx=(0,6))
        ttk.Button(row_btn, text="🧪 Тест токена", command=self.test_token).pack(side="left")

        # Большая кнопка Запуск
        self.start_btn=tk.Button(left, text="▶  ЗАПУСТИТЬ ПАРСИНГ", command=self.start, bg="#15803d", fg="white", font=("Segoe UI", 12, "bold"), relief="flat", padx=12, pady=8, cursor="hand2", activebackground="#16a34a")
        self.start_btn.pack(fill="x", pady=(10,0))
        self.stop_btn=ttk.Button(left, text="■ Стоп", command=self.stop, state="disabled")
        self.stop_btn.pack(fill="x", pady=(4,0))
        ttk.Label(left, text="Resume: при повторе продолжит с места (JSONL). Лимит 10к/сутки — переключает токен сам.", font=("Segoe UI",7), foreground="#888").pack(anchor="w", pady=(4,0))

        # Статистика токенов
        ttk.Separator(left).pack(fill="x", pady=8)
        hdr=ttk.Frame(left)
        hdr.pack(fill="x")
        ttk.Label(hdr, text="Расход сегодня (МСК):", font=("Segoe UI", 8, "bold")).pack(side="left")
        ttk.Button(hdr, text="↻", width=3, command=self.refresh_token_stats).pack(side="right")
        self.stats_frame=ttk.Frame(left)
        self.stats_frame.pack(fill="x", pady=4)

        # drag&drop
        try:
            self.drop_target_register('DND_Files')
            self.dnd_bind('<<Drop>>', self.on_drop)
        except: pass
        # fallback: bind file drop via tk (Windows)
        self.bind("<ButtonRelease-1>", lambda e: None)

        # Правая панель
        right=ttk.Frame(paned, padding=6)
        paned.add(right, weight=2)

        self.progress=ttk.Progressbar(right, mode="determinate")
        self.progress.pack(fill="x", pady=(0,6))
        stats_row=ttk.Frame(right)
        stats_row.pack(fill="x")
        self.stat_total=ttk.Label(stats_row, text="Всего: 0", font=("Segoe UI", 9, "bold"))
        self.stat_total.pack(side="left", padx=(0,12))
        self.stat_done=ttk.Label(stats_row, text="Готово: 0")
        self.stat_done.pack(side="left", padx=(0,12))
        self.stat_found=ttk.Label(stats_row, text="Найдено: 0", foreground="#15803d")
        self.stat_found.pack(side="left", padx=(0,12))
        self.stat_remaining=ttk.Label(stats_row, text="Осталось: —")
        self.stat_remaining.pack(side="left", padx=(0,12))
        self.stat_status=ttk.Label(stats_row, text="Готов", foreground="#0B5394")
        self.stat_status.pack(side="right")

        # Вкладки лог/таблица
        nb=ttk.Notebook(right)
        nb.pack(fill="both", expand=True, pady=(6,0))
        tab_log=ttk.Frame(nb, padding=4)
        tab_table=ttk.Frame(nb, padding=4)
        nb.add(tab_log, text="📜 Лог")
        nb.add(tab_table, text="📋 Таблица (последние 50)")

        self.log_text=scrolledtext.ScrolledText(tab_log, font=("Consolas", 8), bg="#0f172a", fg="#e2e8f0", wrap="word")
        self.log_text.pack(fill="both", expand=True)

        cols=("inn","fio","post","status")
        self.tree=ttk.Treeview(tab_table, columns=cols, show="headings", height=18)
        for col, txt, w in [("inn","ИНН",110),("fio","Руководитель",170),("post","Должность",180),("status","Статус",90)]:
            self.tree.heading(col, text=txt)
            self.tree.column(col, width=w, anchor="center" if col in ("inn","status") else "w")
        vs=ttk.Scrollbar(tab_table, orient="vertical", command=self.tree.yview)
        vs.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=vs.set)
        self.tree.pack(fill="both", expand=True, side="left")
        self.tree.bind("<Double-1>", self.copy_inn)

        self.log("Готов. Вставьте токены, прокси и выберите файл → Запустить.")
        self.log("Подсказка: прокси 95.78.126.13:14755:5nqb8nr:9cz93bt уже вставлен для примера.")

    # --- helpers ---
    def show_welcome(self):
        messagebox.showinfo("Первый запуск",
            "Добро пожаловать в DaData Parser!\n\n"
            "1. Вставьте ваш токен DaData (dadata.ru → Профиль → API ключ)\n"
            "   Можно несколько — каждый даёт +10 000 запросов в сутки.\n"
            "2. Прокси уже настроен (95.78... с авторотацией) — менять не нужно.\n"
            "3. Выберите файл с ИНН (.xlsx/.txt/.csv) и нажмите Запустить.\n\n"
            "Вся установка зависимостей проходит сама при запуске.")

    def check_deps_banner(self):
        miss=missing_deps()
        if miss:
            self.deps_banner.pack(fill="x", padx=12, pady=(0,4))
            self.deps_label.config(text=f"⚠ Нет зависимостей: {', '.join(miss)} — нажмите «Установить всё» (30 сек)")
        else:
            self.deps_banner.pack_forget()

    def install_deps(self):
        self.log("Ставлю зависимости ...")
        self.deps_btn.config(state="disabled", text="Ставлю...")
        def do():
            import pathlib
            req=pathlib.Path(__file__).resolve().parent.parent.parent / "requirements.txt"
            if not req.exists():
                req=pathlib.Path(__file__).resolve().parent.parent / "requirements.txt"
            cmd=[sys.executable, "-m", "pip", "install", "-r", str(req)]
            try:
                subprocess.check_call(cmd)
                self.log("✅ Зависимости установлены")
            except subprocess.CalledProcessError as e:
                self.log(f"Пробую --user ...")
                try:
                    subprocess.check_call(cmd + ["--user"])
                    self.log("✅ Установлено (--user)")
                except Exception as e2:
                    self.log(f"❌ Ошибка: {e2}")
            self.after(0, lambda: self.deps_btn.config(state="normal", text="📦 Установить всё (1 клик)"))
            self.after(0, self.check_deps_banner)
        threading.Thread(target=do, daemon=True).start()

    def toggle_proxy(self):
        state="normal" if self.proxy_enabled.get() else "disabled"
        try: self.proxy_text.configure(state=state)
        except: pass

    def log(self, msg):
        self.log_queue.put(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")

    def poll_log(self):
        try:
            while True:
                msg=self.log_queue.get_nowait()
                self.log_text.insert("end", msg+"\n")
                self.log_text.see("end")
        except: pass
        self.after(120, self.poll_log)

    def browse_input(self):
        p=filedialog.askopenfilename(title="Выберите файл с ИНН", filetypes=[("Все подходящие","*.xlsx *.xls *.txt *.csv *.jsonl"),("Excel","*.xlsx *.xls"),("Текст","*.txt *.csv"),("Все","*.*")])
        if p: self.input_var.set(p)

    def browse_output(self):
        p=filedialog.asksaveasfilename(title="Куда сохранить Excel", defaultextension=".xlsx", filetypes=[("Excel","*.xlsx"),("Все","*.*")], initialfile=pathlib.Path(self.output_var.get()).name)
        if p: self.output_var.set(p)

    def on_drop(self, event):
        # drag-drop
        files=event.data.strip().strip("{}").split()
        if files:
            self.input_var.set(files[0])
            self.log(f"Файл перетащен: {files[0]}")

    def copy_inn(self, event):
        sel=self.tree.selection()
        if sel:
            vals=self.tree.item(sel[0])["values"]
            self.clipboard_clear(); self.clipboard_append(vals[0])
            self.log(f"Скопирован ИНН {vals[0]}")

    def save_config(self):
        tokens=[t.strip() for t in self.tokens_text.get("1.0","end").splitlines() if t.strip() and not t.strip().startswith("#")]
        proxies=[p.strip() for p in self.proxy_text.get("1.0","end").splitlines() if p.strip() and not p.strip().startswith("#")]
        # убрать placeholder
        tokens=[t for t in tokens if "второй токен" not in t.lower()]
        proxies=[p for p in proxies if p]
        self.cfg["dadata"]["tokens"]=tokens
        self.cfg["dadata"]["daily_limit_per_token"]=int(self.limit_var.get())
        self.cfg["dadata"]["delay_between_requests"]=float(self.delay_var.get())
        self.cfg["proxy"]["enabled"]=bool(self.proxy_enabled.get())
        self.cfg["proxy"]["list"]=proxies
        self.cfg["output"]["excel_input"]=self.input_var.get().strip() or None
        self.cfg["output"]["excel_output"]=self.output_var.get().strip()
        from ..config import save_config
        save_config(self.cfg)
        messagebox.showinfo("Сохранено", f"✅ Сохранено: {len(tokens)} токенов, {len(proxies)} прокси\n→ {pathlib.Path('config.yaml').resolve()}")
        self.log(f"💾 config.yaml: {len(tokens)} токенов, {len(proxies)} прокси")
        self.refresh_token_stats()

    def refresh_token_stats(self):
        for w in self.stats_frame.winfo_children():
            w.destroy()
        try:
            from ..token_manager import TokenManager
            tokens=[t.strip() for t in self.tokens_text.get("1.0","end").splitlines() if t.strip() and not t.strip().startswith("#") and "второй токен" not in t.lower()]
            if not tokens:
                tokens=self.cfg["dadata"].get("tokens",[])
            if not tokens:
                ttk.Label(self.stats_frame, text="Вставьте токен выше", font=("Segoe UI",7), foreground="#888").pack(anchor="w")
                return
            tm=TokenManager(tokens, daily_limit=int(self.limit_var.get()), db_path=self.cfg["storage"]["sqlite_path"])
            stats=tm.get_stats()
            used,limit=tm.get_total_today()
            ttk.Label(self.stats_frame, text=f"Всего за сегодня: {used}/{limit} ({used/limit*100:.1f}%)" if limit else "", font=("Segoe UI",8,"bold")).pack(anchor="w")
            for s in stats:
                pct=s["count"]/s["limit_per_day"]*100 if s["limit_per_day"] else 0
                col="#16a34a" if pct<80 else "#eab308" if pct<95 else "#dc2626"
                row=ttk.Frame(self.stats_frame)
                row.pack(fill="x", pady=1)
                bar_len=int(pct/100*12)
                bar="█"*bar_len + "░"*(12-bar_len)
                ttk.Label(row, text=f"{bar} {s['token_masked']}: {s['count']}/{s['limit_per_day']} ({pct:.0f}%)", font=("Consolas",7), foreground=col).pack(side="left")
            self.stat_remaining.config(text=f"Осталось: {tm.remaining()}")
        except Exception as e:
            ttk.Label(self.stats_frame, text=f"—", font=("Segoe UI",7), foreground="#888").pack(anchor="w")
        self.after(5000, self.refresh_token_stats)

    def test_proxy(self):
        self.log("🔍 Тест прокси...")
        def do():
            from ..proxy_manager import ProxyManager
            from ..token_manager import TokenManager
            from ..dadata_client import DadataClient
            pm=ProxyManager([p.strip() for p in self.proxy_text.get("1.0","end").splitlines() if p.strip() and not p.strip().startswith("#") and "95.78" not in p or "95.78" in p])
            # фильтр placeholder уже учтён, но оставим как есть
            raw=[p.strip() for p in self.proxy_text.get("1.0","end").splitlines() if p.strip()]
            pm=ProxyManager(raw)
            tokens=[t.strip() for t in self.tokens_text.get("1.0","end").splitlines() if t.strip() and not t.strip().startswith("#") and "второй токен" not in t.lower()]
            if not tokens:
                self.log("❌ Нет токенов для теста прокси")
                return
            tm=TokenManager(tokens, db_path=self.cfg["storage"]["sqlite_path"])
            client=DadataClient(self.cfg, tm, pm if pm.is_enabled() and self.proxy_enabled.get() else None)
            ok, msg=client.test_proxy()
            self.log(f"{'✅' if ok else '❌'} Прокси: {msg}")
        threading.Thread(target=do, daemon=True).start()

    def test_token(self):
        self.log("🔍 Тест токена...")
        def do():
            from ..proxy_manager import ProxyManager
            from ..token_manager import TokenManager
            from ..dadata_client import DadataClient
            tokens=[t.strip() for t in self.tokens_text.get("1.0","end").splitlines() if t.strip() and not t.strip().startswith("#") and "второй токен" not in t.lower()]
            if not tokens:
                self.log("❌ Вставьте токен DaData")
                return
            raw=[p.strip() for p in self.proxy_text.get("1.0","end").splitlines() if p.strip()]
            pm=ProxyManager(raw)
            tm=TokenManager(tokens, db_path=self.cfg["storage"]["sqlite_path"])
            cfg=self.cfg.copy()
            client=DadataClient(cfg, tm, pm if self.proxy_enabled.get() and pm.is_enabled() else None)
            for tok in tokens[:2]:
                ok, msg=client.test_token(tok)
                self.log(f"{'✅' if ok else '❌'} Токен {tok[:6]}...{tok[-4:]}: {msg}")
        threading.Thread(target=do, daemon=True).start()

    def start(self):
        if self.is_running: return
        self.save_config()
        input_path=self.input_var.get().strip()
        if not input_path or not pathlib.Path(input_path).exists():
            messagebox.showwarning("Файл", "Выберите файл с ИНН:\n• Excel с колонкой ИНН\n• или .txt/.csv (каждый ИНН с новой строки)")
            return
        tokens=[t.strip() for t in self.tokens_text.get("1.0","end").splitlines() if t.strip() and not t.strip().startswith("#") and "второй токен" not in t.lower()]
        if not tokens:
            messagebox.showwarning("Токен", "Вставьте хотя бы один токен DaData")
            return
        self.is_running=True
        self.start_btn.config(state="disabled", bg="#6b7280", text="⏳ Работает...")
        self.stop_btn.config(state="normal")
        for i in self.tree.get_children(): self.tree.delete(i)
        self.log_text.delete("1.0","end")
        self.log(f"▶ Старт: {input_path}")
        self.worker_thread=threading.Thread(target=self.worker, daemon=True)
        self.worker_thread.start()

    def stop(self):
        self.is_running=False
        self.log("⏸ Останавливаю после текущего запроса...")
        self.stat_status.config(text="Останавливается...")

    def worker(self):
        from ..proxy_manager import ProxyManager
        from ..token_manager import TokenManager
        from ..dadata_client import DadataClient
        from ..storage import Storage
        self.cfg["proxy"]["enabled"]=bool(self.proxy_enabled.get())
        self.cfg["dadata"]["daily_limit_per_token"]=int(self.limit_var.get())
        self.cfg["dadata"]["delay_between_requests"]=float(self.delay_var.get())
        raw_p=[p.strip() for p in self.proxy_text.get("1.0","end").splitlines() if p.strip()]
        raw_t=[t.strip() for t in self.tokens_text.get("1.0","end").splitlines() if t.strip() and not t.strip().startswith("#") and "второй токен" not in t.lower()]
        pm=ProxyManager(raw_p)
        tm=TokenManager(raw_t, daily_limit=int(self.limit_var.get()), db_path=self.cfg["storage"]["sqlite_path"])
        storage=Storage(self.cfg["storage"]["raw_jsonl"], self.cfg["storage"]["sqlite_path"], self.cfg["storage"]["log_path"], self.cfg["storage"].get("daily_token_log"))
        client=DadataClient(self.cfg, tm, pm if pm.is_enabled() and self.proxy_enabled.get() else None)
        input_path=self.input_var.get().strip()
        output_path=self.output_var.get().strip()
        from ..worker import load_inns_from_input
        try:
            inns=load_inns_from_input(input_path)
        except Exception as e:
            self.log(f"❌ Ошибка чтения ИНН: {e}")
            self.finish(False); return
        self._total=len(inns); self._done=0; self._found=0
        self.after(0, lambda: self.progress.config(maximum=len(inns), value=0))
        self.after(0, lambda: self.stat_total.config(text=f"Всего: {len(inns)}"))

        import time as time_mod
        orig_sleep=time_mod.sleep
        def sleep_wrap(sec):
            for _ in range(int(sec/0.1)+1):
                if not self.is_running: raise KeyboardInterrupt("stop")
                orig_sleep(min(0.1, sec))

        try:
            existing=storage.load_existing_inns()
            to_query=[i for i in inns if i not in existing]
            self.log(f"К запросу: {len(to_query)} (уже готово {len(existing)} — resume)")
            if not to_query:
                self.log("Всё уже запрошено — собираю Excel")
                if input_path.lower().endswith((".xlsx",".xls")) and output_path:
                    from ..worker import enrich_excel
                    import pathlib as _pl
                    try:
                        out=str(_pl.Path(output_path).resolve())
                        self.log(f"📊 Собираю Excel: {input_path} -> {out}")
                        real_out=enrich_excel(input_path, output_path, storage.load_all_records())
                        self.log(f"✅ Excel сохранён: {real_out}")
                        self.log(f"   Откройте: {real_out}")
                    except Exception as e:
                        import traceback as _tb
                        self.log(f"❌ Ошибка создания Excel: {e}")
                        self.log(_tb.format_exc())
                        # fallback CSV
                        try:
                            from ..worker import save_fallback_csv
                            csv_path=_pl.Path(output_path).with_suffix(".csv")
                            save_fallback_csv(storage.load_all_records(), csv_path)
                            self.log(f"⚠️ Создан CSV fallback: {csv_path.resolve()}")
                        except Exception as ee:
                            self.log(f"CSV fallback тоже не удался: {ee}")
                self.finish(True); return
            done=0
            for inn in to_query:
                if not self.is_running:
                    self.log("⏹ Остановлено пользователем")
                    break
                resp, token=client.query_one(inn)
                if resp.get("http_status")==200:
                    tm.record_usage(token, inn=inn, http_status=200, note="ok")
                storage.append_record(inn, token[:6]+"..." if token else "—", resp)
                def upd(inn=inn, resp=resp, token=token):
                    self._done+=1
                    if resp.get("http_status")==200 and resp.get("data",{}).get("suggestions"):
                        self._found+=1
                        d=resp["data"]["suggestions"][0].get("data",{})
                        mgmt=d.get("management",{}) or {}
                        fio=mgmt.get("name",""); post=mgmt.get("post","")
                        if len(self.tree.get_children())>=50:
                            self.tree.delete(self.tree.get_children()[0])
                        self.tree.insert("", "end", values=(inn, fio, post, "найдено"))
                    self.progress["value"]=self._done
                    self.stat_done.config(text=f"Готово: {self._done}")
                    self.stat_found.config(text=f"Найдено: {self._found}")
                    self.stat_remaining.config(text=f"Осталось: {tm.remaining()}")
                    self.stat_status.config(text=tm.format_status()[:80])
                self.after(0, upd)
                if (done+1)%50==0:
                    self.log(f"⋯ {done+1}/{len(to_query)} | {tm.format_status()}")
                done+=1
                if client.delay>0: sleep_wrap(client.delay)
            if input_path.lower().endswith((".xlsx",".xls")) and output_path:
                self.log("📊 Собираю Excel...")
                from ..worker import enrich_excel
                import pathlib as _pl
                try:
                    out=str(_pl.Path(output_path).resolve())
                    self.log(f"   {input_path} -> {out}")
                    real_out=enrich_excel(input_path, output_path, storage.load_all_records())
                    self.log(f"✅ Готово: {real_out}")
                    self.log(f"   Папка: {_pl.Path(real_out).parent}")
                    # автопроверка
                    if _pl.Path(real_out).exists():
                        sz=_pl.Path(real_out).stat().st_size/1024/1024
                        self.log(f"   Размер: {sz:.2f} МБ — откройте в Excel")
                    else:
                        self.log(f"❌ Файл не найден после сохранения! Проверьте права к папке { _pl.Path(real_out).parent }")
                except Exception as e:
                    import traceback as _tb
                    self.log(f"❌ Ошибка создания Excel: {e}")
                    self.log(_tb.format_exc())
                    self.log("⚠️ Пробую сохранить CSV из тех же данных...")
                    try:
                        from ..worker import save_fallback_csv
                        csv_path=_pl.Path(output_path).with_suffix(".csv")
                        save_fallback_csv(storage.load_all_records(), csv_path)
                        self.log(f"✅ CSV fallback сохранён: {csv_path.resolve()} ({csv_path.stat().st_size/1024:.1f} КБ)")
                        messagebox.showwarning("Excel не создался", f"Excel не создался: {e}\n\nНо CSV с теми же данными сохранён:\n{csv_path.resolve()}\n\nТакже все данные в JSONL:\n{storage.raw_jsonl.resolve()}\n\nЗапустите 'python -m dadata_parser --input ...' для пересборки.")
                    except Exception as ee:
                        self.log(f"❌ CSV fallback тоже failed: {ee}")
                        self.log(_tb.format_exc())
                self.log(f"   Лог: {self.cfg['storage']['log_path']}")
                self.log(f"   JSONL: {self.cfg['storage']['raw_jsonl']}")
            self.log(f"🏁 Завершено: {done} done, {self._found} найдено")
            self.finish(True)
        except KeyboardInterrupt:
            self.log("⏹ Прервано")
            self.finish(False)
        except Exception as e:
            if "AllTokensExhausted" in type(e).__name__ or "исчерпаны" in str(e):
                self.log(f"⛔ Лимит исчерпан: {e}")
                self.log("Следующий слот в 00:00 МСК. Добавьте новый токен и запустите снова.")
            else:
                self.log(f"❌ Ошибка: {e}")
                import traceback; self.log(traceback.format_exc())
            self.finish(False)

    def finish(self, success):
        def do():
            self.is_running=False
            self.start_btn.config(state="normal", bg="#15803d", text="▶  ЗАПУСТИТЬ ПАРСИНГ" if not success else "✅ Готово — запустить снова")
            self.stop_btn.config(state="disabled")
            if success:
                messagebox.showinfo("Готово", "Парсинг завершён!\n\nЛог и JSONL сохранены в папке data/")
            self.refresh_token_stats()
        self.after(0, do)

def main():
    # High DPI
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except: pass
    app=App()
    app.mainloop()

if __name__=="__main__":
    main()
