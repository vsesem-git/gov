"""
Tkinter GUI для Dadata Parser — настройки токенов, прокси, логи расхода.
"""
from __future__ import annotations
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pathlib
import threading
import queue
import webbrowser
from datetime import datetime
import sys
import os

try:
    import yaml
except:
    yaml=None

def load_cfg():
    from ..config import load_config
    return load_config()

def save_cfg(cfg, path=None):
    from ..config import save_config
    save_config(cfg, path)

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Dadata Parser — прокси + ротация токенов")
        self.geometry("1100x800")
        self.minsize(1000, 700)
        self.cfg=load_cfg()
        self.log_queue=queue.Queue()
        self.is_running=False
        self.create_widgets()
        self.after(100, self.poll_log)
        self.refresh_token_stats()

    def create_widgets(self):
        style=ttk.Style()
        try: style.theme_use("vista")
        except: style.theme_use("clam")
        style.configure("Header.TLabel", font=("Segoe UI", 14, "bold"), foreground="#0B5394")
        style.configure("Sub.TLabel", font=("Segoe UI", 8), foreground="#666")

        header=ttk.Frame(self, padding=(16,12))
        header.pack(fill="x")
        ttk.Label(header, text="🔍 Dadata Parser — 10k/сутки/токен + SOCKS5", style="Header.TLabel").pack(anchor="w")
        ttk.Label(header, text="Формат прокси: IP:PORT:LOGIN:PASS (пример: 95.78.126.13:14755:5nqb8nr:9cz93bt) • Авторотация провайдера — менять не нужно", style="Sub.TLabel").pack(anchor="w")
        ttk.Label(header, text="API: POST suggestions.dadata.ru/suggestions/api/4_1/rs/findById/party  {query: inn}", font=("Consolas", 8), foreground="#1F4E78").pack(anchor="w", pady=(4,0))

        paned=ttk.PanedWindow(self, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=12, pady=8)

        # Левая панель: настройки
        left=ttk.LabelFrame(paned, text=" Настройки ", padding=10)
        paned.add(left, weight=1)

        # Токены
        ttk.Label(left, text="Токены DaData (каждый до 10 000/сутки):").pack(anchor="w")
        self.tokens_text=scrolledtext.ScrolledText(left, height=4, font=("Consolas", 9))
        self.tokens_text.pack(fill="x", pady=(4,8))
        for t in self.cfg.get("dadata",{}).get("tokens",[]):
            self.tokens_text.insert("end", t+"\n")
        ttk.Label(left, text="Один токен = 10 000 запросов/сутки. При исчерпании — автопереключение на следующий.", font=("Segoe UI",7), foreground="#888").pack(anchor="w")

        # Лимит
        row_lim=ttk.Frame(left)
        row_lim.pack(fill="x", pady=6)
        ttk.Label(row_lim, text="Лимит/сутки/токен:").pack(side="left")
        self.limit_var=tk.IntVar(value=self.cfg["dadata"].get("daily_limit_per_token",10000))
        ttk.Spinbox(row_lim, from_=100, to=50000, increment=100, textvariable=self.limit_var, width=10).pack(side="left", padx=(6,0))
        ttk.Label(row_lim, text="запросов", foreground="#888").pack(side="left", padx=(4,0))

        # Прокси
        ttk.Separator(left).pack(fill="x", pady=8)
        self.proxy_enabled=tk.BooleanVar(value=self.cfg.get("proxy",{}).get("enabled",True))
        ttk.Checkbutton(left, text="Включить SOCKS5 прокси", variable=self.proxy_enabled).pack(anchor="w")
        ttk.Label(left, text="Прокси (IP:PORT:LOGIN:PASS — каждый с новой строки):").pack(anchor="w", pady=(6,0))
        self.proxy_text=scrolledtext.ScrolledText(left, height=3, font=("Consolas", 9))
        self.proxy_text.pack(fill="x", pady=(4,8))
        for p in self.cfg.get("proxy",{}).get("list",[]):
            self.proxy_text.insert("end", p+"\n")
        ttk.Label(left, text="Пул может быть 1 прокси с авторотацией.", font=("Segoe UI",7), foreground="#888").pack(anchor="w")

        # Вход/выход
        ttk.Separator(left).pack(fill="x", pady=8)
        row_in=ttk.Frame(left)
        row_in.pack(fill="x", pady=4)
        ttk.Label(row_in, text="Входной файл:").pack(side="left")
        self.input_var=tk.StringVar(value=self.cfg.get("output",{}).get("excel_input") or "")
        ttk.Entry(row_in, textvariable=self.input_var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row_in, text="Обзор", width=8, command=self.browse_input).pack(side="left")

        row_out=ttk.Frame(left)
        row_out.pack(fill="x", pady=4)
        ttk.Label(row_out, text="Excel выход:").pack(side="left")
        self.output_var=tk.StringVar(value=self.cfg.get("output",{}).get("excel_output") or "./data/enriched.xlsx")
        ttk.Entry(row_out, textvariable=self.output_var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row_out, text="Обзор", width=8, command=self.browse_output).pack(side="left")

        row_delay=ttk.Frame(left)
        row_delay.pack(fill="x", pady=6)
        ttk.Label(row_delay, text="Пауза, сек:").pack(side="left")
        self.delay_var=tk.DoubleVar(value=self.cfg["dadata"].get("delay_between_requests",0.2))
        ttk.Spinbox(row_delay, from_=0.05, to=2.0, increment=0.05, textvariable=self.delay_var, width=8, format="%.2f").pack(side="left", padx=6)
        ttk.Label(row_delay, text="0.2 = 5 r/s", foreground="#888").pack(side="left")

        # Кнопки проверки
        row_test=ttk.Frame(left)
        row_test.pack(fill="x", pady=8)
        ttk.Button(row_test, text="💾 Сохранить config.yaml", command=self.save_config).pack(side="left", padx=(0,6))
        ttk.Button(row_test, text="🧪 Тест прокси", command=self.test_proxy).pack(side="left", padx=(0,6))
        ttk.Button(row_test, text="🧪 Тест токена", command=self.test_token).pack(side="left")

        # Управление
        row_run=ttk.Frame(left)
        row_run.pack(fill="x", pady=(12,0))
        self.start_btn=ttk.Button(row_run, text="▶ Запустить", command=self.start)
        self.start_btn.pack(side="left", fill="x", expand=True, padx=(0,6))
        self.stop_btn=ttk.Button(row_run, text="■ Стоп", command=self.stop, state="disabled")
        self.stop_btn.pack(side="left")

        ttk.Label(left, text="Resume: при повторном запуске продолжит с места (JSONL)", font=("Segoe UI",7), foreground="#888").pack(anchor="w", pady=(6,0))

        # Статистика токенов
        ttk.Separator(left).pack(fill="x", pady=8)
        ttk.Label(left, text="Расход токенов сегодня (МСК):").pack(anchor="w")
        self.stats_frame=ttk.Frame(left)
        self.stats_frame.pack(fill="x", pady=4)
        self.stats_labels=[]

        # Правая панель: лог
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
        self.stat_status=ttk.Label(stats_row, text="Готов", foreground="#0B5394")
        self.stat_status.pack(side="right")

        self.log_text=scrolledtext.ScrolledText(right, font=("Consolas", 8), bg="#0f172a", fg="#e2e8f0", wrap="word")
        self.log_text.pack(fill="both", expand=True, pady=(6,0))

        # Таблица
        cols=("inn","fio","post","status")
        self.tree=ttk.Treeview(right, columns=cols, show="headings", height=10)
        for col, txt, w in [("inn","ИНН",110),("fio","Руководитель",160),("post","Должность",160),("status","Статус",80)]:
            self.tree.heading(col, text=txt)
            self.tree.column(col, width=w, anchor="center" if col in ("inn","status") else "w")
        self.tree.pack(fill="x", pady=(8,0))

    def log(self, msg):
        from datetime import datetime
        self.log_queue.put(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")

    def poll_log(self):
        try:
            while True:
                msg=self.log_queue.get_nowait()
                self.log_text.insert("end", msg+"\n")
                self.log_text.see("end")
        except:
            pass
        self.after(100, self.poll_log)

    def browse_input(self):
        p=filedialog.askopenfilename(filetypes=[("Excel/CSV/TXT","*.xlsx *.xls *.txt *.csv *.jsonl"),("All","*.*")])
        if p: self.input_var.set(p)

    def browse_output(self):
        p=filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")])
        if p: self.output_var.set(p)

    def save_config(self):
        tokens=[t.strip() for t in self.tokens_text.get("1.0","end").splitlines() if t.strip()]
        proxies=[p.strip() for p in self.proxy_text.get("1.0","end").splitlines() if p.strip()]
        self.cfg["dadata"]["tokens"]=tokens
        self.cfg["dadata"]["daily_limit_per_token"]=int(self.limit_var.get())
        self.cfg["dadata"]["delay_between_requests"]=float(self.delay_var.get())
        self.cfg["proxy"]["enabled"]=bool(self.proxy_enabled.get())
        self.cfg["proxy"]["list"]=proxies
        self.cfg["output"]["excel_input"]=self.input_var.get().strip() or None
        self.cfg["output"]["excel_output"]=self.output_var.get().strip()
        from ..config import save_config
        save_config(self.cfg)
        messagebox.showinfo("Сохранено", f"config.yaml сохранён: {len(tokens)} токенов, {len(proxies)} прокси")
        self.log(f"config.yaml сохранён: tokens={len(tokens)} proxies={len(proxies)}")

    def refresh_token_stats(self):
        # перерисовать статистику
        for w in self.stats_frame.winfo_children():
            w.destroy()
        try:
            from ..token_manager import TokenManager
            tm=TokenManager(self.cfg["dadata"]["tokens"], daily_limit=int(self.limit_var.get()), db_path=self.cfg["storage"]["sqlite_path"])
            stats=tm.get_stats()
            used,limit=tm.get_total_today()
            ttk.Label(self.stats_frame, text=f"Всего: {used}/{limit} ({used/limit*100:.1f}% )" if limit else "", font=("Segoe UI",9,"bold")).pack(anchor="w")
            for s in stats:
                pct=s["count"]/s["limit_per_day"]*100 if s["limit_per_day"] else 0
                bar_color="#16a34a" if pct<80 else "#eab308" if pct<95 else "#dc2626"
                row=ttk.Frame(self.stats_frame)
                row.pack(fill="x", pady=1)
                ttk.Label(row, text=f"{s['token_masked']}: {s['count']}/{s['limit_per_day']} ({pct:.1f}%)", font=("Consolas",8), foreground=bar_color).pack(side="left")
        except Exception as e:
            ttk.Label(self.stats_frame, text=f"Нет данных: {e}", font=("Segoe UI",7), foreground="#888").pack(anchor="w")
        self.after(5000, self.refresh_token_stats)

    def test_proxy(self):
        self.log("Тест прокси...")
        def do():
            from ..proxy_manager import ProxyManager
            from ..token_manager import TokenManager
            from ..dadata_client import DadataClient
            pm=ProxyManager([p.strip() for p in self.proxy_text.get("1.0","end").splitlines() if p.strip()])
            tm=TokenManager([t.strip() for t in self.tokens_text.get("1.0","end").splitlines() if t.strip()], db_path=self.cfg["storage"]["sqlite_path"])
            client=DadataClient(self.cfg, tm, pm if pm.is_enabled() else None)
            ok, msg=client.test_proxy()
            self.log(f"Прокси: {'OK' if ok else 'FAIL'} — {msg}")
        threading.Thread(target=do, daemon=True).start()

    def test_token(self):
        self.log("Тест токена...")
        def do():
            from ..proxy_manager import ProxyManager
            from ..token_manager import TokenManager
            from ..dadata_client import DadataClient
            tokens=[t.strip() for t in self.tokens_text.get("1.0","end").splitlines() if t.strip()]
            if not tokens:
                self.log("Нет токенов")
                return
            pm=ProxyManager([p.strip() for p in self.proxy_text.get("1.0","end").splitlines() if p.strip()])
            tm=TokenManager(tokens, db_path=self.cfg["storage"]["sqlite_path"])
            client=DadataClient(self.cfg, tm, pm if self.proxy_enabled.get() and pm.is_enabled() else None)
            for tok in tokens[:1]:
                ok, msg=client.test_token(tok)
                self.log(f"Токен {tok[:6]}...: {'OK' if ok else 'FAIL'} — {msg}")
        threading.Thread(target=do, daemon=True).start()

    def start(self):
        if self.is_running: return
        # сохранить конфиг
        self.save_config()
        input_path=self.input_var.get().strip()
        if not input_path or not pathlib.Path(input_path).exists():
            messagebox.showwarning("Входной файл", "Выберите существующий входной файл (Excel/TXT с ИНН)")
            return
        self.is_running=True
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        for i in self.tree.get_children(): self.tree.delete(i)
        self.log_text.delete("1.0","end")
        self.log(f"Старт: {input_path}")
        # запуск в потоке
        self.worker_thread=threading.Thread(target=self.worker, daemon=True)
        self.worker_thread.start()

    def stop(self):
        self.is_running=False
        self.log("Остановка после текущего запроса...")
        self.stat_status.config(text="Останавливается...")

    def worker(self):
        from ..proxy_manager import ProxyManager
        from ..token_manager import TokenManager
        from ..dadata_client import DadataClient
        from ..storage import Storage
        from ..worker import Worker
        self.cfg["proxy"]["enabled"]=bool(self.proxy_enabled.get())
        self.cfg["dadata"]["daily_limit_per_token"]=int(self.limit_var.get())
        pm=ProxyManager([p.strip() for p in self.proxy_text.get("1.0","end").splitlines() if p.strip()])
        tm=TokenManager([t.strip() for t in self.tokens_text.get("1.0","end").splitlines() if t.strip()], daily_limit=int(self.limit_var.get()), db_path=self.cfg["storage"]["sqlite_path"])
        storage=Storage(self.cfg["storage"]["raw_jsonl"], self.cfg["storage"]["sqlite_path"], self.cfg["storage"]["log_path"], self.cfg["storage"].get("daily_token_log"))
        client=DadataClient(self.cfg, tm, pm if pm.is_enabled() else None)
        worker=Worker(self.cfg, client, storage, tm)
        input_path=self.input_var.get().strip()
        output_path=self.output_var.get().strip()
        # загрузка ИНН для прогресса
        from ..worker import load_inns_from_input
        try:
            inns=load_inns_from_input(input_path)
        except Exception as e:
            self.log(f"Ошибка загрузки ИНН: {e}")
            self.finish(False)
            return
        self._total=len(inns)
        self._done=0; self._found=0
        self.after(0, lambda: self.progress.config(maximum=len(inns)))

        # Monkey-patch sleep to respect stop
        import time as time_mod
        orig_sleep=time_mod.sleep
        def sleep_wrap(sec):
            # прерываем если остановлено
            for _ in range(int(sec/0.1)+1):
                if not self.is_running: raise KeyboardInterrupt("stop")
                orig_sleep(min(0.1, sec))

        # Переопределим client delay? Пусть worker сам вызывает, но прервём через исключение
        try:
            # Ручной batch чтобы обновлять UI
            from ..token_manager import AllTokensExhausted
            existing=storage.load_existing_inns()
            to_query=[i for i in inns if i not in existing]
            self.log(f"К запросу: {len(to_query)} (уже {len(existing)} в JSONL)")
            done=0
            for inn in to_query:
                if not self.is_running:
                    self.log("Остановлено пользователем")
                    break
                resp, token=client.query_one(inn)
                # учёт
                if resp.get("http_status")==200:
                    tm.record_usage(token, inn=inn, http_status=200, note="ok")
                storage.append_record(inn, token[:6]+"..." if token else "—", resp)
                # UI
                def update_ui(inn=inn, resp=resp):
                    self._done+=1
                    if resp.get("http_status")==200 and resp.get("data",{}).get("suggestions"):
                        self._found+=1
                        d=resp["data"]["suggestions"][0].get("data",{})
                        mgmt=d.get("management",{}) or {}
                        fio=mgmt.get("name","")
                        post=mgmt.get("post","")
                        if len(self.tree.get_children())>=50:
                            self.tree.delete(self.tree.get_children()[0])
                        self.tree.insert("", "end", values=(inn, fio, post, "найдено"))
                    self.progress["value"]=self._done
                    self.stat_done.config(text=f"Готово: {self._done}")
                    self.stat_found.config(text=f"Найдено: {self._found}")
                    self.stat_status.config(text=f"{tm.format_status()}")
                self.after(0, update_ui)
                # прогресс-лог
                if (done+1)%50==0:
                    self.log(f"{done+1}/{len(to_query)} {tm.format_status()}")
                done+=1
                if client.delay>0:
                    # проверяемый sleep
                    sleep_wrap(client.delay)
            # финал Excel
            if input_path.lower().endswith((".xlsx",".xls")) and output_path:
                self.log("Генерация Excel...")
                from ..worker import enrich_excel
                results=storage.load_all_records()
                enrich_excel(input_path, output_path, results)
                self.log(f"Excel: {output_path}")
            self.log(f"Готово: done={done} found={self._found}")
            self.finish(True)
        except KeyboardInterrupt:
            self.log("Прервано")
            self.finish(False)
        except AllTokensExhausted as e:
            self.log(f"Лимит исчерпан: {e}")
            self.finish(False)
        except Exception as e:
            self.log(f"Ошибка: {e}")
            import traceback
            self.log(traceback.format_exc())
            self.finish(False)

    def finish(self, success):
        def do():
            self.is_running=False
            self.start_btn.config(state="normal")
            self.stop_btn.config(state="disabled")
            if success:
                self.start_btn.config(text="✅ Готово — запустить снова")
                messagebox.showinfo("Готово", "Парсинг завершён")
            else:
                self.start_btn.config(text="▶ Запустить снова")
        self.after(0, do)

def main():
    app=App()
    app.mainloop()

if __name__=="__main__":
    main()
