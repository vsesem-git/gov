"""
ProxyManager — SOCKS5 прокси IP:PORT:LOGIN:PASS → socks5:// URL.

Поддерживает форматы:
  95.78.126.13:14755:5nqb8nr:9cz93bt   (IP:PORT:LOGIN:PASS — 4 части через :)
  95.78.126.13:14755;5nqb8nr;9cz93bt   (через ;)
  95.78.126.13:14755                  (без авторизации)

Авторотация провайдера: достаточно 1 прокси в списке, он сам меняет IP.
При пуле — round_robin / random / sticky.
"""
from __future__ import annotations
import random
import re
from urllib.parse import quote
from typing import List, Optional


class ProxyManager:
    def __init__(self, proxy_list: List[str], rotation: str = "round_robin"):
        self.raw_list = [p.strip() for p in (proxy_list or []) if p.strip()]
        self.rotation = rotation or "round_robin"
        self._idx = 0
        self._sticky = None

    @staticmethod
    def parse_proxy_string(s: str) -> dict:
        """
        Парсит строку прокси в компоненты.
        Возвращает dict: {host, port, login, password, original}
        """
        s = s.strip()
        # Попытка разделить: сначала по ; если есть с логином
        # Формат с ; login:pass уже задан примером IP:PORT:LOGIN:PASS через :
        # Универсально: разбиваем по : и ;
        original = s
        # Заменяем ; на : для унификации, но сохраняем если внутри пароля есть символы
        # Проще: если в строке 4 части через : → то IP:PORT:LOGIN:PASS
        # Если 3 части через : где последняя без цифры? Будем ветвиться
        # Нормализация: заменить ; на :
        norm = s.replace(";", ":")
        parts = norm.split(":")

        if len(parts) == 4:
            # IP:PORT:LOGIN:PASS
            host, port, login, pwd = parts
            return {"host": host.strip(), "port": port.strip(), "login": login.strip(), "password": pwd.strip(), "original": original}
        elif len(parts) == 2:
            # IP:PORT без авторизации
            host, port = parts
            return {"host": host.strip(), "port": port.strip(), "login": None, "password": None, "original": original}
        elif len(parts) == 1:
            # возможно host:port без схемы?
            return {"host": parts[0], "port": "", "login": None, "password": None, "original": original}
        else:
            # Нестандарт: попробуем regex
            # fallback: считать первый : как host:port, остальное login:pass
            # Но для простоты вернём как есть
            host = parts[0]
            port = parts[1] if len(parts) > 1 else ""
            login = parts[2] if len(parts) > 2 else None
            pwd = ":".join(parts[3:]) if len(parts) > 3 else None
            return {"host": host.strip(), "port": port.strip(), "login": login.strip() if login else None, "password": pwd.strip() if pwd else None, "original": original}

    @staticmethod
    def to_socks5_url(parsed: dict) -> str:
        """Преобразует parsed прокси в socks5h:// URL для requests."""
        host = parsed["host"]
        port = parsed["port"]
        login = parsed.get("login")
        pwd = parsed.get("password")
        if login and pwd:
            # URL-encode login/password (на случай спецсимволов)
            login_enc = quote(login, safe="")
            pwd_enc = quote(pwd, safe="")
            return f"socks5h://{login_enc}:{pwd_enc}@{host}:{port}"
        elif login:
            login_enc = quote(login, safe="")
            return f"socks5h://{login_enc}@{host}:{port}"
        else:
            return f"socks5h://{host}:{port}"

    def is_enabled(self) -> bool:
        return len(self.raw_list) > 0

    def get_current(self) -> Optional[dict]:
        if not self.raw_list:
            return None
        raw = self.raw_list[self._idx % len(self.raw_list)]
        return self.parse_proxy_string(raw)

    def get_proxy_dict(self) -> Optional[dict]:
        """Возвращает dict для requests: {'http': url, 'https': url} или None."""
        if not self.raw_list:
            return None
        cur = self.get_current()
        if not cur:
            return None
        url = self.to_socks5_url(cur)
        return {"http": url, "https": url}

    def rotate(self):
        """Переключает на следующий прокси (round_robin)."""
        if not self.raw_list:
            return
        if self.rotation == "random":
            self._idx = random.randint(0, len(self.raw_list) - 1)
        elif self.rotation == "sticky":
            if self._sticky is None:
                self._sticky = random.choice(self.raw_list)
            # sticky — не ротируем
            return
        else:  # round_robin
            self._idx = (self._idx + 1) % len(self.raw_list)

    def get_masked(self) -> str:
        cur = self.get_current()
        if not cur:
            return "— (прямое подключение)"
        host = cur.get("host", "")
        port = cur.get("port", "")
        return f"{host}:{port} (***скрыт***)"

    def list_masked(self) -> List[str]:
        out = []
        for raw in self.raw_list:
            p = self.parse_proxy_string(raw)
            out.append(f"{p['host']}:{p['port']} {'с логином' if p.get('login') else 'без логина'}")
        return out


def proxy_string_to_url(proxy_str: str) -> str:
    """Утилита: строка → socks5h URL."""
    parsed = ProxyManager.parse_proxy_string(proxy_str)
    return ProxyManager.to_socks5_url(parsed)
