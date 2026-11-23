"""
Dop OKVED provider — дозагрузка дополнительных ОКВЭДов.
DaData findById/party возвращает только основной okved, поле okveds всегда null.
Для каталога нужен второй ОКВЭД → тянем из ФНС ЕГРЮЛ.

Провайдеры:
- EgrulNalogProvider: https://egrul.nalog.ru (бесплатно, без токена, ~1 r/s)
- StubProvider: заглушка для оффлайн-тестов

Использование:
    provider = EgrulNalogProvider(delay=1.0, proxy_manager=pm)
    dop = provider.fetch_dop_okveds(inn)  # -> ["36.00.2", "42.21", ...]

Интеграция в Worker: после основного парсинга DaData, если dop_okved.enabled,
проходим по тем же ИНН и дозапрашиваем доп.коды, пишем в Storage как dop_okveds.

Примечание: провайдер ФНС возвращает все ОКВЭДы, первый = основной (совпадает с DaData),
остальные = доп. Мы сохраняем их как массив кодов.
"""
from __future__ import annotations
import time
import json
import re
import logging
from typing import List, Optional, Dict
import requests

logger = logging.getLogger("dadata_parser")

class BaseDopProvider:
    def fetch_dop_okveds(self, inn: str) -> List[str]:
        raise NotImplementedError
    def test(self) -> bool:
        return True

class StubProvider(BaseDopProvider):
    def fetch_dop_okveds(self, inn: str) -> List[str]:
        return []

class EgrulNalogProvider(BaseDopProvider):
    """
    ФНС ЕГРЮЛ: https://egrul.nalog.ru
    Поток:
      POST https://egrul.nalog.ru/  {query: inn}
        -> {"t": "xxxxx..."}  (токен запроса)
      GET https://egrul.nalog.ru/search-result/{t}
        -> {"status":"done", "n": ...} или сразу данные
      GET https://egrul.nalog.ru/vyp/{t}.json  или /n/{t}.json
        -> {"rows": [{"v": {"okved": "68.32.1", "okveds": [...], "inn": "...", "name": ...}}]}

    На практике структура немного меняется, поэтому парсим гибко.
    """
    BASE = "https://egrul.nalog.ru"
    def __init__(self, delay: float = 1.0, timeout: int = 15, proxy_manager=None, verify_ssl: bool = True):
        self.delay = delay
        self.timeout = timeout
        self.pm = proxy_manager
        self.verify_ssl = verify_ssl
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0",
            "Accept": "application/json, text/plain, */*",
        })

    def _proxies(self):
        if self.pm and self.pm.is_enabled():
            return self.pm.get_proxy_dict()
        return None

    def fetch_dop_okveds(self, inn: str) -> List[str]:
        inn = str(inn).strip()
        if not inn: return []
        try:
            # 1. запрос токена
            r = self.session.post(
                f"{self.BASE}/",
                data={"query": inn, "vyp3CaptchaToken": "", "page": "", "region": "", "PreventChromeAutocomplete": ""},
                timeout=self.timeout,
                proxies=self._proxies(),
                verify=self.verify_ssl,
            )
            if r.status_code != 200:
                logger.debug(f"EGRUL POST {inn} HTTP {r.status_code}")
                return []
            j = r.json()
            t = j.get("t")
            if not t:
                # иногда сразу rows
                if "rows" in j:
                    return self._extract_okveds(j)
                return []
            time.sleep(self.delay * 0.5)
            # 2. опрос результата (может потребовать подождать)
            for attempt in range(5):
                r2 = self.session.get(
                    f"{self.BASE}/search-result/{t}",
                    params={"r": str(int(time.time()*1000)), "n": "1"},
                    timeout=self.timeout,
                    proxies=self._proxies(),
                    verify=self.verify_ssl,
                )
                if r2.status_code != 200:
                    time.sleep(0.8)
                    continue
                j2 = r2.json()
                # если статус еще не готов
                if j2.get("status") == "wait":
                    time.sleep(0.8)
                    continue
                # данные могут быть прямо в j2 или надо дернуть vyp
                if "rows" in j2:
                    return self._extract_okveds(j2)
                # попробуем vyp JSON
                try:
                    r3 = self.session.get(
                        f"{self.BASE}/vyp/{t}",
                        timeout=self.timeout,
                        proxies=self._proxies(),
                        verify=self.verify_ssl,
                    )
                    # иногда Content-Type text/html с json внутри
                    if "rows" in r3.text:
                        return self._extract_okveds(r3.json())
                except: pass
                time.sleep(0.5)
            return []
        except Exception as e:
            logger.debug(f"EGRUL dop okved error {inn}: {e}")
            return []
        finally:
            if self.delay:
                time.sleep(self.delay)

    def _extract_okveds(self, data: Dict) -> List[str]:
        # data может быть {"rows": [{"v": {"okved": "...", "okveds": [{"k": "...", "n": "..."}], ...}}]}
        # или {"vyp": {"okved": ...}}
        try:
            rows = data.get("rows") or data.get("vyp", {}).get("rows") or []
            if not rows:
                # прямой путь
                v = data.get("v") or data.get("rows", [{}])[0].get("v") if data.get("rows") else None
                if v:
                    return self._v_to_okveds(v)
                return []
            # берём первую запись по ИНН (обычно 1)
            v = rows[0].get("v") if isinstance(rows[0], dict) else rows[0]
            return self._v_to_okveds(v)
        except Exception as e:
            logger.debug(f"_extract_okveds parse err: {e} data={str(data)[:500]}")
            return []

    def _v_to_okveds(self, v: Dict) -> List[str]:
        # v может содержать okved, okveds, dopOkveds, etc.
        candidates = []
        # основной
        main = v.get("okved") or v.get("okvedOsn") or v.get("okved_2014") or ""
        if main:
            candidates.append(str(main).strip())
        # доп массив — разные ключи
        for key in ["okveds", "dopOkveds", "okvedDop", "okved2", "okveds2", "okveds_2014"]:
            arr = v.get(key)
            if isinstance(arr, list):
                for item in arr:
                    if isinstance(item, dict):
                        code = item.get("k") or item.get("code") or item.get("okved") or item.get("n")
                        if code: candidates.append(str(code).strip())
                    elif isinstance(item, str):
                        candidates.append(item.strip())
            elif isinstance(arr, str) and arr:
                candidates.append(arr.strip())
        # нормализация ОКВЭД кодов: только вида 00.00 или 00.00.0
        okved_re = re.compile(r"^\d{2}(\.\d{1,2}){0,2}$")
        out=[]
        seen=set()
        for c in candidates:
            c=c.strip().replace(",",".")
            if okved_re.match(c) and c not in seen:
                seen.add(c)
                out.append(c)
        # первый уже основной, остальные доп
        return out

    def test(self) -> bool:
        try:
            # тест на известном ИНН сбербанка
            res = self.fetch_dop_okveds("7707083893")
            return isinstance(res, list)
        except: return False

# Фабрика по config
def create_provider(cfg: dict, proxy_manager=None):
    dop_cfg = cfg.get("dop_okved", {}) if cfg else {}
    if not dop_cfg.get("enabled"):
        return StubProvider()
    provider_name = dop_cfg.get("provider", "egrul_nalog")
    if provider_name == "egrul_nalog":
        return EgrulNalogProvider(
            delay=float(dop_cfg.get("delay", 1.0)),
            timeout=int(dop_cfg.get("timeout", 15)),
            proxy_manager=proxy_manager,
            verify_ssl=dop_cfg.get("verify_ssl", True),
        )
    else:
        return StubProvider()
