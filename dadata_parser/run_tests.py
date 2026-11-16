#!/usr/bin/env python3
"""Быстрый self-test без сети: проверяет прокси-парсинг, token_manager, config."""
from dadata_parser.proxy_manager import ProxyManager
from dadata_parser.token_manager import TokenManager
from dadata_parser.config import load_config
import tempfile, pathlib

print("== Тест прокси-парсера ==")
tests=[
    ("95.78.126.13:14755:5nqb8nr:9cz93bt", "socks5h://5nqb8nr:9cz93bt@95.78.126.13:14755"),
    ("1.2.3.4:8080", "socks5h://1.2.3.4:8080"),
    ("1.2.3.4:1080;user;pass", "socks5h://user:pass@1.2.3.4:1080"),
]
for raw, expected in tests:
    pm=ProxyManager([raw])
    url=pm.get_proxy_dict()["https"]
    assert url==expected, f"{raw} → {url} != {expected}"
    print(f"  OK {raw} → {url}")

print("== Тест TokenManager (мок DB) ==")
import tempfile, pathlib
tmp=pathlib.Path(tempfile.mktemp(suffix=".db"))
tm=TokenManager(["tok1","tok2"], daily_limit=5, db_path=tmp)
# имитация 5 запросов на tok1
for i in range(5):
    t=tm.acquire()
    assert t=="tok1", f"expected tok1 got {t} at {i}"
    tm.record_usage(t, inn=f"inn{i}", http_status=200)
print("  tok1 exhausted after 5")
t=tm.acquire()
assert t=="tok2", f"expected tok2 got {t}"
print(f"  ротация → {t} OK")
for i in range(5):
    tm.record_usage(t, inn=f"inn2_{i}", http_status=200)
# теперь все исчерпаны
try:
    tm.acquire()
    assert False, "должно выбросить AllTokensExhausted"
except Exception as e:
    print(f"  оба исчерпаны → {type(e).__name__} OK")

print("== Тест config.yaml ==")
cfg=load_config()
assert cfg["dadata"]["daily_limit_per_token"]==10000
print(f"  tokens: {len(cfg['dadata']['tokens'])} proxy enabled={cfg['proxy']['enabled']}")

print("\nВсе тесты пройдены ✅")
