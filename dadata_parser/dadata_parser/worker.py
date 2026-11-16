"""
Worker — оркестратор: загрузка ИНН, resume, пакетный парсинг, обогащение Excel.
"""
from __future__ import annotations
import pathlib
import json
import time
import logging
from datetime import datetime
from typing import List, Optional

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    import yaml
except:
    yaml = None

from .config import masked_token
from .token_manager import AllTokensExhausted
from .storage import Storage

try:
    from tqdm import tqdm
    HAS_TQDM = True
except ImportError:
    HAS_TQDM = False

logger = logging.getLogger("dadata_parser")

# DaData → новые колонки (совместимо с enrich_dadata.py)
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
NEW_HEADER_FILL = PatternFill(start_color="548235", end_color="548235", fill_type="solid")
THIN_BORDER = Border(left=Side(style='thin', color='B4C6E7'), right=Side(style='thin', color='B4C6E7'), top=Side(style='thin', color='B4C6E7'), bottom=Side(style='thin', color='B4C6E7'))
ALT_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=10, name="Calibri")

def normalize_site(site):
    if not site: return ""
    s=str(site).strip().lower()
    for p in ["https://","http://","www."]:
        if s.startswith(p):
            s=s[len(p):]
    return s.rstrip("/").strip()

def normalize_email(email):
    if not email: return ""
    return str(email).strip().lower()

def extract_fields(dadata_resp):
    if not dadata_resp or dadata_resp.get("http_status") != 200 or not dadata_resp.get("data"):
        return {
            "fio": "", "post": "", "management_type": "", "email": "", "site": "", "phone": "",
            "ogrn": "", "full_name": "", "status": "", "actuality": "",
            "raw_compact": json.dumps(dadata_resp, ensure_ascii=False)[:32000] if dadata_resp else ""
        }
    data_wrapper=dadata_resp["data"]
    sugg=data_wrapper.get("suggestions") or []
    if not sugg:
        return {"fio":"", "post":"", "management_type":"", "email":"", "site":"", "phone":"", "ogrn":"", "full_name":"", "status":"Не найдено в ЕГРЮЛ", "actuality":"", "raw_compact": json.dumps(dadata_resp, ensure_ascii=False)[:32000]}
    d=sugg[0].get("data") or {}
    mgmt=d.get("management") or {}
    fio=str(mgmt.get("name") or "").strip()
    post=str(mgmt.get("post") or "").strip()
    mtype=str(mgmt.get("type") or "").strip()
    email=str(d.get("email") or "").strip()
    site=str(d.get("site") or "").strip()
    phone=str(d.get("phone") or "").strip()
    ogrn=str(d.get("ogrn") or "").strip()
    full_name=str(sugg[0].get("unrestricted_value") or sugg[0].get("value") or "").strip()
    state=d.get("state") or {}
    status=str(state.get("status") or "").strip()
    actuality=str(state.get("actuality_date") or state.get("registration_date") or "").strip()
    return {"fio":fio,"post":post,"management_type":mtype,"email":email,"site":site,"phone":phone,"ogrn":ogrn,"full_name":full_name,"status":status,"actuality":actuality,"raw_compact": json.dumps(dadata_resp, ensure_ascii=False)[:32000]}


def load_inns_from_input(input_path: str | pathlib.Path) -> List[str]:
    p=pathlib.Path(input_path)
    if not p.exists():
        raise FileNotFoundError(f"Входной файл не найден: {p}")
    ext=p.suffix.lower()
    inns=[]
    if ext in (".txt",".csv",".tsv"):
        text=p.read_text(encoding="utf-8", errors="ignore")
        import re
        inns=re.findall(r"\b\d{10}\b|\b\d{12}\b", text)
        # uniq preserve
        seen=set()
        uniq=[]
        for i in inns:
            if i not in seen:
                seen.add(i)
                uniq.append(i)
        return uniq
    elif ext in (".xlsx",".xls"):
        wb=openpyxl.load_workbook(p, read_only=True, data_only=True)
        seen=set()
        for ws in wb.worksheets:
            if ws.title=="Оглавление":
                continue
            rows=list(ws.iter_rows(values_only=True))
            if not rows: continue
            header=rows[0]
            try:
                inn_idx=header.index("ИНН")
            except ValueError:
                inn_idx=5
            for row in rows[1:]:
                if len(row) <= inn_idx: continue
                val=row[inn_idx]
                if val is None: continue
                s=str(val).strip().replace(" ","")
                if s.isdigit() and len(s) in (10,12) and s not in seen:
                    seen.add(s)
                    inns.append(s)
        wb.close()
        return inns
    elif ext==".jsonl":
        seen=set()
        with open(p, encoding="utf-8") as f:
            import re
            for line in f:
                for m in re.findall(r"\b\d{10}\b|\b\d{12}\b", line):
                    if m not in seen:
                        seen.add(m)
                        inns.append(m)
        return inns
    else:
        # try as txt
        text=p.read_text(encoding="utf-8", errors="ignore")
        import re
        inns=re.findall(r"\b\d{10}\b|\b\d{12}\b", text)
        return list(dict.fromkeys(inns))


def enrich_excel(input_path: str | pathlib.Path, output_path: str | pathlib.Path, results: dict[str, dict]):
    """Обогащает Excel новыми колонками (совместимо со старым парсером)."""
    wb=openpyxl.load_workbook(input_path, data_only=True)
    for ws in wb.worksheets:
        if ws.title=="Оглавление":
            continue
        header_len=ws.max_column
        # add headers
        for idx, name in enumerate(NEW_COLUMNS, start=header_len+1):
            c=ws.cell(row=1, column=idx, value=name)
            c.fill=NEW_HEADER_FILL
            c.font=HEADER_FONT
            c.alignment=Alignment(horizontal='center', vertical='center', wrap_text=True)
            c.border=THIN_BORDER
        for cell in ws[1]:
            if cell.column <= header_len:
                cell.fill=HEADER_FILL
                cell.font=HEADER_FONT
                cell.alignment=Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.border=THIN_BORDER
        ws.row_dimensions[1].height=38
        ws.freeze_panes="A2"
        ws.auto_filter.ref=ws.dimensions
        ws.sheet_properties.pageSetUpPr.fitToPage=True
        ws.page_setup.orientation=ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize=ws.PAPERSIZE_A3
        ws.page_setup.fitToWidth=1
        widths=[55,30,30,32,16,14,13,32,18,26,22,32,16,18, 26,28,18,26,26,18,18,18,18,16,26,28,16,16,32]
        for i,w in enumerate(widths, start=1):
            if i <= ws.max_column:
                ws.column_dimensions[get_column_letter(i)].width=w
        for r_idx in range(2, ws.max_row+1):
            inn=str(ws.cell(row=r_idx, column=6).value or "").strip()  # F=6 ИНН
            email_reg=str(ws.cell(row=r_idx, column=10).value or "").strip()
            site_reg=str(ws.cell(row=r_idx, column=8).value or "").strip()
            resp=results.get(inn, {})
            fields=extract_fields(resp)
            fio=fields["fio"]; post=fields["post"]; mtype=fields["management_type"]
            email_dd=fields["email"]; site_dd=fields["site"]; phone_dd=fields["phone"]
            ogrn_dd=fields["ogrn"]; full_dd=fields["full_name"]; status_dd=fields["status"]; actuality_dd=fields["actuality"]; raw_compact=fields["raw_compact"]
            email_match="Да" if email_dd and normalize_email(email_dd)==normalize_email(email_reg) else ("Нет" if email_dd and email_reg else ("Пусто в DaData" if not email_dd else "Пусто в реестре" if not email_reg else "Нет"))
            site_match="Да" if site_dd and normalize_site(site_dd)==normalize_site(site_reg) else ("Нет" if site_dd and site_reg else ("Пусто в DaData" if not site_dd else "Пусто в реестре" if not site_reg else "Нет"))
            has_extra_email="Да" if email_dd and normalize_email(email_dd)!=normalize_email(email_reg) and email_dd else "Нет"
            has_extra_site="Да" if site_dd and normalize_site(site_dd)!=normalize_site(site_reg) and site_dd else "Нет"
            if resp.get("http_status")==200 and resp.get("data") and resp["data"].get("suggestions"):
                req_status="Найдено"
            elif resp.get("http_status")==200:
                req_status="Не найдено"
            elif resp.get("http_status"):
                req_status=f"Ошибка {resp.get('http_status')}"
            else:
                req_status=resp.get("error","Не запрошено")[:80]
            queried_at=datetime.now().strftime("%Y-%m-%d %H:%M")
            values=[fio, post, mtype, email_dd, site_dd, phone_dd, email_match, site_match, has_extra_email, has_extra_site, ogrn_dd, full_dd, status_dd, actuality_dd, req_status, queried_at, raw_compact]
            for off,val in enumerate(values, start=0):
                ws.cell(row=r_idx, column=header_len+1+off, value=val)
        for r_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            is_alt=(r_idx%2==0)
            for cell in row:
                cell.alignment=Alignment(vertical='center', wrap_text=True)
                cell.border=THIN_BORDER
                cell.font=Font(size=9, name="Calibri")
                if is_alt:
                    cell.fill=ALT_FILL
            ws.row_dimensions[r_idx].height=16
    pathlib.Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


class Worker:
    def __init__(self, config: dict, dadata_client, storage: Storage, token_manager):
        self.cfg=config
        self.client=dadata_client
        self.storage=storage
        self.tm=token_manager
        self.logger=logging.getLogger("dadata_parser")

    def run(self, input_path: str | pathlib.Path | None, inns: List[str] | None = None, output_path: str | pathlib.Path | None = None, enable_tqdm: bool = True) -> dict:
        """
        Основной прогон.
        - input_path: путь к Excel/txt/csv (если None — берём inns напрямую)
        - inns: список ИНН (приоритет если задан)
        """
        # загрузка ИНН
        if inns is None:
            if not input_path:
                raise ValueError("Нужно указать --input или список inns")
            inns=load_inns_from_input(input_path)
        else:
            inns=[str(i).strip() for i in inns if str(i).strip()]
            # uniq preserve order
            seen=set(); uniq=[]
            for i in inns:
                if i not in seen:
                    seen.add(i); uniq.append(i)
            inns=uniq

        total=len(inns)
        if total==0:
            self.logger.warning("ИНН не найдено во входном файле")
            return {"total":0, "done":0, "found":0}

        self.logger.info(f"Найдено ИНН: {total} | вход: {input_path} | прокси: {self.client.pm.get_masked() if self.client.pm and self.client.pm.is_enabled() else 'прямое'}")
        self.logger.info(f"Токены: {self.tm.format_status()} | лимит/сутки/токен: {self.tm.daily_limit}")

        existing=self.storage.load_existing_inns()
        to_query=[i for i in inns if i not in existing]
        self.logger.info(f"Resume: уже запрошено {len(existing)} | к запросу {len(to_query)}")

        if not to_query:
            self.logger.info("Все ИНН уже запрошены — собираю Excel")
            results=self.storage.load_all_records()
            if output_path and str(input_path).lower().endswith((".xlsx",".xls")):
                enrich_excel(input_path, output_path, results)
                self.logger.info(f"Excel обновлён: {output_path}")
            return {"total": total, "done": len(existing), "found": sum(1 for r in results.values() if r.get("http_status")==200 and r.get("data",{}).get("suggestions"))}

        # прогресс
        pbar=None
        if enable_tqdm and HAS_TQDM:
            pbar=tqdm(total=len(to_query), desc="DaData", unit=" ИНН")

        done=0; found=0; errors=0
        every_n=self.cfg.get("logging",{}).get("console_every_n", 50)

        def on_each(inn, resp, token):
            nonlocal done, found, errors
            done+=1
            if resp.get("http_status")==200 and resp.get("data",{}).get("suggestions"):
                found+=1
            elif resp.get("http_status")!=200:
                errors+=1
            # запись JSONL
            token_masked=masked_token(token)
            self.storage.append_record(inn, token_masked, resp, queried_at=datetime.now().isoformat())
            if pbar:
                pbar.update(1)
                pbar.set_postfix_str(f"{token_masked} {done}/{len(to_query)} found={found}")
            # лог каждые N
            if done % every_n == 0:
                used,limit=self.tm.get_total_today()
                stats=self.tm.get_stats()
                self.logger.info(f"Прогресс {done}/{len(to_query)} found={found} err={errors} | {self.tm.format_status()} | всего {used}/{limit}")
                self.storage.log_token_daily(stats, used, limit)

        try:
            self.client.batch(to_query, sleep=True, on_each=on_each)
        except AllTokensExhausted as e:
            self.logger.error(str(e))
            if pbar: pbar.close()
            # сохраняем ежедневный лог
            stats=self.tm.get_stats()
            used,limit=self.tm.get_total_today()
            self.storage.log_token_daily(stats, used, limit)
            # всё равно пытаемся сохранить частичный результат
            results=self.storage.load_all_records()
            if output_path and input_path and str(input_path).lower().endswith((".xlsx",".xls")):
                enrich_excel(input_path, output_path, results)
            raise
        finally:
            if pbar and not pbar.disable:
                pbar.close()

        # финал
        stats=self.tm.get_stats()
        used,limit=self.tm.get_total_today()
        self.logger.info(f"Готово: done={done} found={found} errors={errors} | {self.tm.format_status()} | всего за сутки {used}/{limit}")
        self.storage.log_token_daily(stats, used, limit)

        # Excel
        if output_path and input_path and str(input_path).lower().endswith((".xlsx",".xls")):
            results=self.storage.load_all_records()
            enrich_excel(input_path, output_path, results)
            self.logger.info(f"Excel сохранён: {output_path}")

        return {"total": total, "done": done, "found": found, "errors": errors, "remaining": self.tm.remaining()}

