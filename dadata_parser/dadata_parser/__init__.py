"""
Dadata Parser — парсер ЕГРЮЛ по ИНН через DaData Find-Party
с SOCKS5 прокси и ротацией токенов (10k/сутки/токен).
"""
__version__ = "1.0.0"
__all__ = ["config", "token_manager", "proxy_manager", "dadata_client", "storage", "worker"]
