"""
Загрузка конфигурации: config.yaml + ENV + CLI overrides.
"""
from __future__ import annotations
import os
import pathlib
import yaml
from typing import List, Optional

DEFAULT_CONFIG_PATH = pathlib.Path(__file__).parent.parent / "config.yaml"
EXAMPLE_CONFIG_PATH = pathlib.Path(__file__).parent.parent / "config.yaml.example"


def load_config(path: Optional[str | pathlib.Path] = None) -> dict:
    """Загружает config.yaml, применяет ENV."""
    cfg_path = pathlib.Path(path) if path else DEFAULT_CONFIG_PATH
    if not cfg_path.exists():
        cfg_path = EXAMPLE_CONFIG_PATH
        if not cfg_path.exists():
            return _default_config()

    with open(cfg_path, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f) or {}

    # ENV overrides
    # DADATA_TOKENS=token1,token2
    env_tokens = os.getenv("DADATA_TOKENS") or os.getenv("DADATA_TOKEN")
    if env_tokens:
        tokens = [t.strip() for t in env_tokens.split(",") if t.strip()]
        if tokens:
            cfg.setdefault("dadata", {})["tokens"] = tokens

    env_proxy = os.getenv("DADATA_PROXY")
    if env_proxy:
        cfg.setdefault("proxy", {})["list"] = [p.strip() for p in env_proxy.split(",") if p.strip()]

    # Нормализация
    cfg = _normalize(cfg)
    return cfg


def _normalize(cfg: dict) -> dict:
    cfg.setdefault("dadata", {})
    cfg["dadata"].setdefault("tokens", [])
    cfg["dadata"].setdefault("daily_limit_per_token", 10000)
    cfg["dadata"].setdefault("timeout", 15)
    cfg["dadata"].setdefault("retries", 3)
    cfg["dadata"].setdefault("delay_between_requests", 0.2)
    cfg["dadata"].setdefault("retry_backoff_base", 1.5)
    cfg["dadata"].setdefault("handle_429_pause", 2.0)

    cfg.setdefault("proxy", {})
    cfg["proxy"].setdefault("enabled", False)
    cfg["proxy"].setdefault("list", [])
    cfg["proxy"].setdefault("rotation", "round_robin")
    cfg["proxy"].setdefault("verify_ssl", True)

    cfg.setdefault("storage", {})
    cfg["storage"].setdefault("dir", "./data")
    cfg["storage"].setdefault("sqlite_path", "./data/parser.db")
    cfg["storage"].setdefault("raw_jsonl", "./data/dadata_raw.jsonl")
    cfg["storage"].setdefault("log_path", "./data/parser.log")
    cfg["storage"].setdefault("daily_token_log", "./data/token_usage_{date}.log")

    cfg.setdefault("output", {})
    cfg["output"].setdefault("excel_input", None)
    cfg["output"].setdefault("excel_output", "./data/enriched.xlsx")
    cfg["output"].setdefault("save_raw_jsonl", True)

    cfg.setdefault("dop_okved", {})
    cfg["dop_okved"].setdefault("enabled", False)
    cfg["dop_okved"].setdefault("provider", "egrul_nalog")
    cfg["dop_okved"].setdefault("delay", 1.0)
    cfg["dop_okved"].setdefault("timeout", 15)
    cfg["dop_okved"].setdefault("verify_ssl", True)

    cfg.setdefault("logging", {})
    cfg["logging"].setdefault("level", "INFO")
    cfg["logging"].setdefault("console_every_n", 50)
    cfg["logging"].setdefault("max_log_file_mb", 20)

    # фильтруем пустые токены
    cfg["dadata"]["tokens"] = [t.strip() for t in cfg["dadata"]["tokens"] if t and t.strip()]
    cfg["proxy"]["list"] = [p.strip() for p in cfg["proxy"]["list"] if p and p.strip()]

    return cfg


def _default_config() -> dict:
    return _normalize({})


def masked_token(t: str) -> str:
    if not t:
        return "—"
    t = t.strip()
    if len(t) <= 10:
        return t[:3] + "..."
    return t[:6] + "..." + t[-4:]


def save_config(cfg: dict, path: Optional[str | pathlib.Path] = None):
    p = pathlib.Path(path) if path else DEFAULT_CONFIG_PATH
    p.parent.mkdir(parents=True, exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        yaml.safe_dump(cfg, f, allow_unicode=True, sort_keys=False)
