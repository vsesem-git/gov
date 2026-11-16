"""
Storage — JSONL + SQLite для resume и логов.
"""
from __future__ import annotations
import json
import pathlib
import sqlite3
import logging
from datetime import datetime
from typing import Iterable, Set

try:
    import zoneinfo
    TZ_MSK = zoneinfo.ZoneInfo("Europe/Moscow")
except ImportError:
    import datetime as dt
    TZ_MSK = dt.timezone(dt.timedelta(hours=3))


class Storage:
    def __init__(self, raw_jsonl: str | pathlib.Path, sqlite_path: str | pathlib.Path, log_path: str | pathlib.Path, daily_token_log_pattern: str | None = None):
        self.raw_jsonl = pathlib.Path(raw_jsonl)
        self.sqlite_path = pathlib.Path(sqlite_path)
        self.log_path = pathlib.Path(log_path)
        self.daily_token_log_pattern = daily_token_log_pattern
        # ensure dirs
        for p in [self.raw_jsonl, self.sqlite_path, self.log_path]:
            if str(p) and p.parent:
                p.parent.mkdir(parents=True, exist_ok=True)
        # logger
        self.logger = self._setup_logger()

    def _setup_logger(self) -> logging.Logger:
        logger = logging.getLogger("dadata_parser")
        logger.setLevel(logging.INFO)
        logger.propagate = False
        if not logger.handlers:
            # file handler
            fh = logging.FileHandler(self.log_path, encoding="utf-8")
            fh.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s", datefmt="%Y-%m-%d %H:%M:%S"))
            logger.addHandler(fh)
            # also console via StreamHandler added in worker
        return logger

    def load_existing_inns(self) -> Set[str]:
        """Загружает уже запрошенные ИНН из JSONL для resume."""
        if not self.raw_jsonl.exists():
            return set()
        inns = set()
        with open(self.raw_jsonl, "r", encoding="utf-8") as f:
            for line in f:
                line=line.strip()
                if not line:
                    continue
                try:
                    obj=json.loads(line)
                    inn=obj.get("inn")
                    if inn:
                        inns.add(str(inn).strip())
                except:
                    continue
        return inns

    def append_record(self, inn: str, token_masked: str, response: dict, queried_at: str | None = None):
        if queried_at is None:
            queried_at = datetime.now(TZ_MSK).isoformat()
        record = {
            "inn": str(inn).strip(),
            "queried_at": queried_at,
            "token_masked": token_masked,
            "response": response
        }
        with open(self.raw_jsonl, "a", encoding="utf-8") as f:
            f.write(json.dumps(record, ensure_ascii=False) + "\n")

    def load_all_records(self) -> dict[str, dict]:
        """Возвращает dict inn -> response (последний)."""
        if not self.raw_jsonl.exists():
            return {}
        out = {}
        with open(self.raw_jsonl, "r", encoding="utf-8") as f:
            for line in f:
                line=line.strip()
                if not line:
                    continue
                try:
                    obj=json.loads(line)
                    out[str(obj["inn"]).strip()] = obj["response"]
                except:
                    continue
        return out

    def log_token_daily(self, stats: list[dict], total_used: int, total_limit: int):
        """Пишет ежедневный файл token_usage_{date}.log"""
        if not self.daily_token_log_pattern:
            return
        today = datetime.now(TZ_MSK).strftime("%Y-%m-%d")
        path = pathlib.Path(self.daily_token_log_pattern.format(date=today))
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now(TZ_MSK).isoformat()}] Всего: {total_used}/{total_limit} ({total_used/total_limit*100:.1f}% если total_limit>0 else 0)\n")
            for row in stats:
                f.write(f"  {row['token_masked']}: {row['count']}/{row['limit_per_day']} ({row['count']/row['limit_per_day']*100:.1f}%) updated={row['updated_at']}\n")
            f.write("\n")

    def iter_raw_records(self):
        if not self.raw_jsonl.exists():
            return
        with open(self.raw_jsonl, "r", encoding="utf-8") as f:
            for line in f:
                line=line.strip()
                if not line:
                    continue
                try:
                    yield json.loads(line)
                except:
                    continue
