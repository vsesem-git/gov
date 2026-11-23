"""
Task Manager — экспорт/импорт задания для переноса между версиями парсера.
Задание = набор ИНН + конфиг + прогресс (JSONL cursor).

Формат .dtask.json:
{
  "version": "1.0",
  "created_at": "2026-11-23T12:00:00+03:00",
  "source_file": "/path/input.xlsx",
  "output_file": "/path/output.xlsx",
  "config": { ... snapshot ... },
  "inns": ["1234567890", ...],
  "inns_total": 36108,
  "progress": {"done": 12000, "remaining": 24108, "jsonl_rows": 12000},
  "token_stats": [...]
}
Также умеет паковать в .zip с вложенным data/dadata_raw.jsonl и parser.db.
"""
from __future__ import annotations
import json, pathlib, datetime, zipfile
from typing import List, Optional

try:
    import zoneinfo
    TZ_MSK = zoneinfo.ZoneInfo("Europe/Moscow")
except ImportError:
    import datetime as dt
    TZ_MSK = dt.timezone(dt.timedelta(hours=3))

class TaskManager:
    def __init__(self, config: dict = None, storage=None, token_manager=None):
        # поддерживать оба порядка: (config, storage) и (storage, tm, config)
        if isinstance(config, dict) and storage is not None and hasattr(storage, 'raw_jsonl'):
            # новый: TaskManager(cfg, storage)
            self.cfg = config
            self.storage = storage
            self.tm = token_manager
        elif config is not None and hasattr(config, 'raw_jsonl'):
            # старый: TaskManager(storage, tm, cfg)
            self.storage = config
            self.tm = storage
            self.cfg = token_manager or {}
        else:
            self.cfg = config or {}
            self.storage = storage
            self.tm = token_manager

    def build_task(self, inns: List[str] | None, source_file: str | None = None, output_file: str | None = None) -> dict:
        inns = inns or []
        done_inns = self.storage.load_existing_inns() if self.storage else set()
        done = len(done_inns & set(inns)) if inns else len(done_inns)
        remaining = len(inns) - done if inns else 0
        progress = {"done": done, "remaining": remaining, "jsonl_rows": len(done_inns), "done_inns_sample": list(done_inns)[:10]}
        token_stats=[]
        if self.tm and hasattr(self.tm, 'get_stats'):
            try: token_stats=self.tm.get_stats()
            except: pass
        return {
            "version": "1.0",
            "created_at": datetime.datetime.now(TZ_MSK).isoformat(),
            "source_file": str(source_file) if source_file else "",
            "output_file": str(output_file) if output_file else "",
            "config": self.cfg,
            "inns": inns,
            "inns_total": len(inns),
            "progress": progress,
            "token_stats": token_stats,
        }

    def export_task_json(self, dest_path: str | pathlib.Path, inns: List[str] | None = None, source_file: str | None = None, output_file: str | None = None, **kwargs) -> pathlib.Path:
        # поддержка старого вызова export_task_json(inns, source, output, dest)
        if isinstance(dest_path, list):
            # старая сигнатура (inns, source, output, dest_path)
            inns_old, source_old, output_old, dest_old = dest_path, inns, source_file, output_file
            return self.export_task_json(dest_old, inns_old, source_old, output_old)
        dest = pathlib.Path(dest_path)
        dest.parent.mkdir(parents=True, exist_ok=True)
        task = self.build_task(inns, source_file, output_file)
        with open(dest, "w", encoding="utf-8") as f:
            json.dump(task, f, ensure_ascii=False, indent=2)
        return dest

    def export_task_zip(self, dest_path: str | pathlib.Path, inns: List[str] | None = None, source_file: str | None = None, output_file: str | None = None, **kwargs) -> pathlib.Path:
        if isinstance(dest_path, list):
            inns_old, source_old, output_old, dest_old = dest_path, inns, source_file, output_file
            return self.export_task_zip(dest_old, inns_old, source_old, output_old)
        dest = pathlib.Path(dest_path)
        dest.parent.mkdir(parents=True, exist_ok=True)
        # json внутри zip
        task = self.build_task(inns, source_file, output_file)
        import tempfile, os
        with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("task.json", json.dumps(task, ensure_ascii=False, indent=2))
            if self.storage:
                for p in [getattr(self.storage, 'raw_jsonl', None), getattr(self.storage, 'sqlite_path', None)]:
                    if not p: continue
                    pp = pathlib.Path(p)
                    if pp.exists():
                        z.write(pp, arcname=f"data/{pp.name}")
            if source_file and pathlib.Path(source_file).exists():
                try:
                    if pathlib.Path(source_file).stat().st_size < 10*1024*1024:
                        z.write(source_file, arcname=f"input/{pathlib.Path(source_file).name}")
                except: pass
        return dest

    @staticmethod
    def import_task_json(path: str | pathlib.Path) -> dict:
        with open(path, encoding="utf-8") as f:
            return json.load(f)

    def import_task_zip(self, path: str | pathlib.Path) -> dict:
        # распаковывает data/* в storage dir
        p = pathlib.Path(path)
        with zipfile.ZipFile(p, "r") as z:
            # читаем task.json
            try:
                data = z.read("task.json")
                task = json.loads(data.decode("utf-8"))
            except:
                task = {}
            # извлекаем data/* в storage.dir
            if self.storage:
                target_dir = pathlib.Path(getattr(self.storage, 'raw_jsonl', 'data/dadata_raw.jsonl')).parent
                target_dir.mkdir(parents=True, exist_ok=True)
                for name in z.namelist():
                    if name.startswith("data/"):
                        out = target_dir / pathlib.Path(name).name
                        # не перезаписываем если JSONL уже есть и больше — merge
                        content = z.read(name)
                        if out.exists() and name.endswith(".jsonl"):
                            # append недостающие? пока просто оставим существующий, а новый допишем в .imported
                            # проще: если файл есть — дописываем новые строки если их нет
                            # сейчас просто перезапишем если импортируемый больше
                            try:
                                if out.stat().st_size >= len(content):
                                    continue
                            except: pass
                        out.write_bytes(content)
            return task

    def apply_task(self, task: dict) -> tuple[List[str], dict]:
        inns = task.get("inns") or []
        cfg = task.get("config") or self.cfg
        return inns, cfg
