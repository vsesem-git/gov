"""
TokenManager — ротация токенов DaData с лимитом 10 000/сутки/токен.

Каждый токен имеет дневной лимит (по умолчанию 10000). После исчерпания
переключается на следующий. Если все исчерпаны — AllTokensExhausted.

Хранение:
- SQLite таблица token_usage (date, token_hash, token_masked, count, limit)
- In-memory кэш + запись при каждом инкременте (можно батчить)

Дата — по Europe/Moscow (где бьёт DaData), но храним и UTC.
"""
from __future__ import annotations
import hashlib
import sqlite3
import pathlib
import datetime
from typing import List, Optional, Tuple
import threading

from .config import masked_token

try:
    import zoneinfo
    TZ_MSK = zoneinfo.ZoneInfo("Europe/Moscow")
except ImportError:
    # Python 3.11: zoneinfo есть, fallback
    import datetime as dt
    TZ_MSK = datetime.timezone(datetime.timedelta(hours=3))


class AllTokensExhausted(Exception):
    pass


class TokenManager:
    def __init__(self, tokens: List[str], daily_limit: int = 10000, db_path: str | pathlib.Path = "./data/parser.db"):
        self.tokens = [t.strip() for t in tokens if t.strip()]
        if not self.tokens:
            raise ValueError("Не задано ни одного токена DaData")
        self.daily_limit = int(daily_limit)
        self.db_path = pathlib.Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.Lock()
        self._init_db()
        # кэш на сегодня: token -> count
        self._today_counts: dict[str, int] = {}
        self._load_today_counts()
        # указатель текущего токена
        self._current_idx = 0
        # подобрать первый не исчерпанный
        self._pick_next_available()

    def _today_msk(self) -> str:
        now = datetime.datetime.now(TZ_MSK)
        return now.strftime("%Y-%m-%d")

    def _token_hash(self, token: str) -> str:
        return hashlib.sha256(token.encode()).hexdigest()[:12]

    def _init_db(self):
        with sqlite3.connect(self.db_path) as con:
            con.execute("""
                CREATE TABLE IF NOT EXISTS token_usage (
                    date TEXT NOT NULL,
                    token_hash TEXT NOT NULL,
                    token_masked TEXT NOT NULL,
                    count INTEGER NOT NULL DEFAULT 0,
                    limit_per_day INTEGER NOT NULL,
                    updated_at TEXT NOT NULL,
                    PRIMARY KEY (date, token_hash)
                )
            """)
            # для логов по часам (опционально)
            con.execute("""
                CREATE TABLE IF NOT EXISTS token_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    ts TEXT NOT NULL,
                    date TEXT NOT NULL,
                    token_hash TEXT NOT NULL,
                    token_masked TEXT NOT NULL,
                    inn TEXT,
                    http_status INTEGER,
                    note TEXT
                )
            """)
            con.commit()

    def _load_today_counts(self):
        today = self._today_msk()
        with sqlite3.connect(self.db_path) as con:
            con.row_factory = sqlite3.Row
            for token in self.tokens:
                th = self._token_hash(token)
                row = con.execute("SELECT count FROM token_usage WHERE date=? AND token_hash=?", (today, th)).fetchone()
                self._today_counts[token] = row["count"] if row else 0

    def get_current_token(self) -> Optional[str]:
        """Возвращает текущий доступный токен или None если все исчерпаны."""
        with self._lock:
            # найти первый подходящий начиная с current_idx
            for offset in range(len(self.tokens)):
                idx = (self._current_idx + offset) % len(self.tokens)
                tok = self.tokens[idx]
                cnt = self._today_counts.get(tok, 0)
                if cnt < self.daily_limit:
                    self._current_idx = idx
                    return tok
            return None

    def _pick_next_available(self):
        tok = self.get_current_token()
        # если None — все исчерпаны, оставим индекс 0 (для ошибки)
        if tok is None:
            self._current_idx = 0

    def acquire(self) -> str:
        """Выдаёт токен для следующего запроса. Бросает AllTokensExhausted если лимит исчерпан."""
        tok = self.get_current_token()
        if tok is None:
            raise AllTokensExhausted(
                f"Все токены исчерпаны на сегодня ({self._today_msk()}): {len(self.tokens)} шт × {self.daily_limit} = {len(self.tokens)*self.daily_limit} запросов/сутки. Следующий слот в 00:00 МСК."
            )
        return tok

    def record_usage(self, token: str, inn: str = "", http_status: Optional[int] = None, note: str = ""):
        """Инкремент счётчика для токена (успешный запрос)."""
        with self._lock:
            tok = token.strip()
            today = self._today_msk()
            # in-memory
            self._today_counts[tok] = self._today_counts.get(tok, 0) + 1
            cnt = self._today_counts[tok]
            th = self._token_hash(tok)
            masked = masked_token(tok)
            now_iso = datetime.datetime.now(TZ_MSK).isoformat()

            # SQLite upsert
            with sqlite3.connect(self.db_path) as con:
                con.execute("""
                    INSERT INTO token_usage(date, token_hash, token_masked, count, limit_per_day, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?)
                    ON CONFLICT(date, token_hash) DO UPDATE SET
                        count=excluded.count,
                        updated_at=excluded.updated_at
                """, (today, th, masked, cnt, self.daily_limit, now_iso))
                # лог-по-запросно (можно отключить для производительности)
                con.execute("""
                    INSERT INTO token_log(ts, date, token_hash, token_masked, inn, http_status, note)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (now_iso, today, th, masked, inn, http_status, note))
                con.commit()

            # если исчерпали лимит — переключаем
            if cnt >= self.daily_limit:
                # попробуем переключить на следующий при следующем acquire
                pass

    def mark_exhausted(self, token: str):
        """Принудительно помечает токен как исчерпанный сегодня (например, при 429 с сообщением о лимите)."""
        with self._lock:
            tok = token.strip()
            today = self._today_msk()
            th = self._token_hash(tok)
            masked = masked_token(tok)
            now_iso = datetime.datetime.now(TZ_MSK).isoformat()
            # ставим count = limit
            self._today_counts[tok] = self.daily_limit
            with sqlite3.connect(self.db_path) as con:
                con.execute("""
                    INSERT INTO token_usage(date, token_hash, token_masked, count, limit_per_day, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?)
                    ON CONFLICT(date, token_hash) DO UPDATE SET
                        count=excluded.count,
                        updated_at=excluded.updated_at
                """, (today, th, masked, self.daily_limit, self.daily_limit, now_iso))
                con.commit()

    def get_stats(self, date: Optional[str] = None) -> List[dict]:
        """Возвращает статистику по всем токенам на дату."""
        if date is None:
            date = self._today_msk()
        with sqlite3.connect(self.db_path) as con:
            con.row_factory = sqlite3.Row
            rows = con.execute(
                "SELECT date, token_masked, count, limit_per_day, updated_at FROM token_usage WHERE date=? ORDER BY count DESC",
                (date,)
            ).fetchall()
            # дополнить токены которые ещё не использовались (0)
            existing_hashes = {r["token_masked"] for r in rows}
            out = []
            for r in rows:
                out.append(dict(r))
            # добавить неиспользованные
            for tok in self.tokens:
                m = masked_token(tok)
                if m not in existing_hashes:
                    out.append({"date": date, "token_masked": m, "count": 0, "limit_per_day": self.daily_limit, "updated_at": ""})
            return out

    def get_total_today(self) -> Tuple[int, int]:
        """Возвращает (использованно, лимит всего)."""
        total_used = sum(self._today_counts.get(t, 0) for t in self.tokens)
        total_limit = len(self.tokens) * self.daily_limit
        return total_used, total_limit

    def remaining(self) -> int:
        used, total = self.get_total_today()
        return max(0, total - used)

    def format_status(self) -> str:
        parts = []
        for tok in self.tokens:
            cnt = self._today_counts.get(tok, 0)
            pct = (cnt / self.daily_limit * 100) if self.daily_limit else 0
            m = masked_token(tok)
            parts.append(f"{m}: {cnt}/{self.daily_limit} ({pct:.1f}%)")
        return " | ".join(parts)

    def force_reload(self):
        with self._lock:
            self._load_today_counts()
            self._pick_next_available()
