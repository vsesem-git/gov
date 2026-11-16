"""
DadataClient — запросы к suggestions.dadata.ru с учётом прокси, ретраев и ротации токенов.

Endpoint: POST https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/party
          Body: {"query": "<inn>"}
          Header: Authorization: Token <key>
"""
from __future__ import annotations
import time
import json
import logging
from typing import Optional, Tuple, List

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from .token_manager import TokenManager, AllTokensExhausted
from .proxy_manager import ProxyManager

DADATA_URL = "https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/party"

logger = logging.getLogger("dadata_parser")


class DadataClient:
    def __init__(self, config: dict, token_manager: TokenManager, proxy_manager: Optional[ProxyManager] = None):
        self.cfg = config
        self.tm = token_manager
        self.pm = proxy_manager
        d_cfg = config.get("dadata", {})
        self.timeout = int(d_cfg.get("timeout", 15))
        self.retries = int(d_cfg.get("retries", 3))
        self.delay = float(d_cfg.get("delay_between_requests", 0.2))
        self.backoff_base = float(d_cfg.get("retry_backoff_base", 1.5))
        self.pause_429 = float(d_cfg.get("handle_429_pause", 2.0))

        # Session с пулом
        self.session = requests.Session()
        # retry на уровне HTTPAdapter (для сетевых сбоев)
        retry = Retry(total=2, backoff_factor=0.5, status_forcelist=[502, 503, 504])
        self.session.mount("https://", HTTPAdapter(max_retries=retry))
        self.session.headers.update({
            "Content-Type": "application/json",
            "Accept": "application/json",
            "User-Agent": "dadata_parser/1.0 (github: vsesem-git/gov)"
        })

    def _proxy_payload(self) -> Optional[dict]:
        if not self.pm or not self.pm.is_enabled():
            return None
        return self.pm.get_proxy_dict()

    def _headers(self, token: str) -> dict:
        return {
            "Content-Type": "application/json",
            "Accept": "application/json",
            "Authorization": f"Token {token}"
        }

    def query_one(self, inn: str) -> Tuple[dict, str]:
        """
        Делает 1 запрос по ИНН с ротацией токенов и ретраями.
        Возвращает (response_dict, used_token).
        response_dict: {inn, http_status, data?, error?, attempt}
        """
        inn = str(inn).strip()
        last_err = None
        # Попытки с учётом смены токена при исчерпании: до len(tokens) * retries
        attempt = 0
        while attempt < self.retries * max(1, len(self.tm.tokens)):
            attempt += 1
            try:
                token = self.tm.acquire()
            except AllTokensExhausted as e:
                # проброс наверх
                raise

            payload = {"query": inn}
            proxies = self._proxy_payload()
            headers = self._headers(token)

            try:
                resp = self.session.post(
                    DADATA_URL,
                    json=payload,
                    headers=headers,
                    proxies=proxies,
                    timeout=self.timeout,
                    verify=self.cfg.get("proxy", {}).get("verify_ssl", True)
                )
            except requests.exceptions.ProxyError as e:
                logger.warning(f"[ProxyError] ИНН {inn} прокси {self.pm.get_masked() if self.pm else '—'}: {e}")
                last_err = f"ProxyError: {e}"
                # ротация прокси и повтор с тем же токеном?
                if self.pm:
                    self.pm.rotate()
                time.sleep(self.backoff_base ** min(attempt, 4))
                continue
            except requests.exceptions.RequestException as e:
                last_err = str(e)
                logger.warning(f"[NetErr] ИНН {inn}: {e} (attempt {attempt})")
                time.sleep(self.backoff_base ** min(attempt, 4))
                continue

            # Обработка статусов
            if resp.status_code == 429:
                # Лимит — считаем токен исчерпанным на сегодня?
                # DaData 429 = дневной лимит или rate limit. Сначала пауза, при повторе — mark exhausted
                body = resp.text[:1000]
                logger.warning(f"[429] ИНН {inn} token {token[:6]}... — пауза {self.pause_429 * attempt:.1f}s | body: {body[:300]}")
                # Если тело содержит "daily" или "limit" — сразу помечаем
                if "limit" in body.lower() or "daily" in body.lower() or attempt >= 2:
                    logger.warning(f"  → помечаем токен {token[:6]}... как исчерпанный на сегодня (429)")
                    self.tm.mark_exhausted(token)
                    # попробовать другой токен без задержки
                    continue
                time.sleep(self.pause_429 * attempt)
                continue

            if resp.status_code == 401:
                # неверный токен — помечаем, пробуем следующий
                logger.error(f"[401] Неверный токен {token[:6]}... для ИНН {inn}: {resp.text[:300]}")
                self.tm.mark_exhausted(token)
                last_err = f"401 Unauthorized: {resp.text[:500]}"
                continue

            if resp.status_code == 403:
                logger.error(f"[403] Доступ запрещён токен {token[:6]}... ИНН {inn}: {resp.text[:500]}")
                self.tm.mark_exhausted(token)
                last_err = f"403 Forbidden: {resp.text[:500]}"
                continue

            if resp.status_code != 200:
                # другие HTTP ошибки — ретрай
                logger.warning(f"[HTTP {resp.status_code}] ИНН {inn}: {resp.text[:500]} (attempt {attempt})")
                last_err = f"HTTP {resp.status_code}: {resp.text[:1000]}"
                time.sleep(self.backoff_base * attempt)
                continue

            # 200 OK
            try:
                data = resp.json()
            except Exception as e:
                last_err = f"JSON decode error: {e} | body {resp.text[:500]}"
                time.sleep(1)
                continue

            # Успех
            return {"inn": inn, "http_status": 200, "data": data, "attempt": attempt}, token

        # исчерпали попытки
        return {"inn": inn, "http_status": None, "error": last_err or "max retries без ответа", "attempt": attempt}, token if 'token' in locals() else ""

    def batch(self, inns: List[str], sleep: bool = True, progress_cb=None, on_each=None):
        """
        Пакетный парсинг.
        - inns: список ИНН
        - on_each: callback(inn, response, token) вызывается после каждого успеха
        - progress_cb: callback(done, total)
        """
        total = len(inns)
        for idx, inn in enumerate(inns, 1):
            try:
                resp, token = self.query_one(inn)
            except AllTokensExhausted as e:
                logger.error(str(e))
                raise  # останов

            # учёт
            # Записываем использование токена только для успешных или для 200/404?
            # Логика: считаем каждый запрос к DaData как 1 из лимита, кроме сетевых ошибок
            # http_status 200 (даже "не найдено") — считается. 401/403 уже помечен.
            if resp.get("http_status") in (200,):
                self.tm.record_usage(token, inn=inn, http_status=200, note="ok")
            elif resp.get("http_status") is None:
                # сетевая ошибка — не считаем, иначе исказим лимит
                pass
            else:
                # HTTP ошибка которая всё равно потратила лимит? считаем консервативно? DaData не списывает при не-200.
                pass

            if on_each:
                on_each(inn, resp, token)
            if progress_cb:
                progress_cb(idx, total)
            if sleep and self.delay > 0:
                time.sleep(self.delay)

    def test_token(self, token: str) -> Tuple[bool, str]:
        """Проверка токена 1 тестовым запросом (ИНН 7707083893 — Сбербанк)."""
        test_inn = "7707083893"
        try:
            proxies = self._proxy_payload()
            resp = self.session.post(
                DADATA_URL,
                json={"query": test_inn},
                headers=self._headers(token),
                proxies=proxies,
                timeout=self.timeout
            )
            if resp.status_code == 200 and resp.json().get("suggestions"):
                return True, f"OK: {len(resp.json()['suggestions'])} suggestions"
            elif resp.status_code == 200:
                return True, "OK: пустой ответ (но токен рабочий)"
            else:
                return False, f"HTTP {resp.status_code}: {resp.text[:500]}"
        except Exception as e:
            return False, str(e)

    def test_proxy(self) -> Tuple[bool, str]:
        """Проверка прокси через httpbin или DaData."""
        if not self.pm or not self.pm.is_enabled():
            return False, "Прокси не настроен"
        # пробуем запросить ip через httpbin
        proxies = self.pm.get_proxy_dict()
        try:
            r = requests.get("https://httpbin.org/ip", proxies=proxies, timeout=10)
            return True, f"Proxy OK: {r.json()}"
        except Exception as e:
            # fallback — пробуем DaData с тестовым токеном
            return False, f"Proxy FAIL: {e} | proxy={self.pm.get_masked()}"
