"""
CLI для Dadata Parser.
Примеры:
  python -m dadata_parser --input inns.txt
  python -m dadata_parser --input реестр.xlsx --excel-output out.xlsx
  python -m dadata_parser --test --inn 5902290635
  python -m dadata_parser --test-proxy
  python -m dadata_parser --gui
"""
from __future__ import annotations
import argparse
import pathlib
import sys
import logging

from .config import load_config, masked_token
from .proxy_manager import ProxyManager
from .token_manager import TokenManager
from .dadata_client import DadataClient
from .storage import Storage
from .worker import Worker

def build_parser():
    p=argparse.ArgumentParser(description="Dadata Parser — ЕГРЮЛ по ИНН (ротация токенов 10k/сутки + SOCKS5 прокси)")
    p.add_argument("--config", default=None, help="Путь к config.yaml (по умолчанию dadata_parser/config.yaml)")
    p.add_argument("--input", dest="input_path", default=None, help="Входной файл: xlsx/txt/csv/jsonl со списком ИНН")
    p.add_argument("--inns", nargs="*", default=None, help="Список ИНН через пробел (альтернатива --input)")
    p.add_argument("--output", dest="output_path", default=None, help="Выходной JSONL (если не Excel)")
    p.add_argument("--excel-output", dest="excel_output", default=None, help="Путь для обогащённого Excel (только если input — Excel)")
    p.add_argument("--tokens", default=None, help="Токены через запятую (перекрывает config.yaml)")
    p.add_argument("--proxy", dest="proxy_str", default=None, help="Прокси IP:PORT:LOGIN:PASS (перекрывает config)")
    p.add_argument("--no-proxy", action="store_true", help="Отключить прокси")
    p.add_argument("--daily-limit", type=int, default=None, help="Лимит/сутки/токен (по умолчанию 10000)")
    p.add_argument("--delay", type=float, default=None, help="Пауза между запросами, сек")
    p.add_argument("--test", action="store_true", help="Тестовый запрос 1 ИНН")
    p.add_argument("--inn", default="5902290635", help="ИНН для --test (по умолчанию Администрация Перми)")
    p.add_argument("--test-proxy", action="store_true", help="Проверить прокси")
    p.add_argument("--gui", action="store_true", help="Запустить GUI")
    p.add_argument("--log-level", default=None, choices=["DEBUG","INFO","WARNING","ERROR"])
    p.add_argument("--no-tqdm", action="store_true", help="Отключить tqdm прогресс-бар")
    return p

def setup_logging(level: str = "INFO", log_path: str | None = None):
    lvl=getattr(logging, level.upper(), logging.INFO)
    root=logging.getLogger()
    root.setLevel(lvl)
    # console
    if not any(isinstance(h, logging.StreamHandler) for h in root.handlers):
        ch=logging.StreamHandler(sys.stdout)
        ch.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s", datefmt="%H:%M:%S"))
        ch.setLevel(lvl)
        root.addHandler(ch)
    if log_path:
        pathlib.Path(log_path).parent.mkdir(parents=True, exist_ok=True)
        fh=logging.FileHandler(log_path, encoding="utf-8")
        fh.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        fh.setLevel(lvl)
        root.addHandler(fh)

def main(argv=None):
    parser=build_parser()
    args=parser.parse_args(argv)

    if args.gui:
        # lazy import GUI (tkinter may be absent)
        try:
            from .gui.app import main as gui_main
            gui_main()
            return 0
        except Exception as e:
            print(f"GUI не запустился: {e}", file=sys.stderr)
            import traceback; traceback.print_exc()
            return 1

    cfg=load_config(args.config)

    # CLI overrides
    if args.tokens:
        cfg["dadata"]["tokens"]=[t.strip() for t in args.tokens.split(",") if t.strip()]
    if args.proxy_str:
        cfg["proxy"]["enabled"]=True
        cfg["proxy"]["list"]=args.proxy_str.split(",")
    if args.no_proxy:
        cfg["proxy"]["enabled"]=False
        cfg["proxy"]["list"]=[]
    if args.daily_limit:
        cfg["dadata"]["daily_limit_per_token"]=args.daily_limit
    if args.delay is not None:
        cfg["dadata"]["delay_between_requests"]=args.delay
    if args.log_level:
        cfg["logging"]["level"]=args.log_level

    setup_logging(cfg["logging"]["level"], cfg["storage"]["log_path"])

    # validate
    if not cfg["dadata"]["tokens"]:
        print("ОШИБКА: не задано ни одного токена DaData (config.yaml → dadata.tokens)", file=sys.stderr)
        return 2

    # managers
    pm=ProxyManager(cfg["proxy"]["list"], rotation=cfg["proxy"].get("rotation","round_robin")) if cfg["proxy"].get("enabled") else None
    tm=TokenManager(cfg["dadata"]["tokens"], daily_limit=cfg["dadata"]["daily_limit_per_token"], db_path=cfg["storage"]["sqlite_path"])
    storage=Storage(cfg["storage"]["raw_jsonl"], cfg["storage"]["sqlite_path"], cfg["storage"]["log_path"], cfg["storage"].get("daily_token_log"))
    client=DadataClient(cfg, tm, pm)
    worker=Worker(cfg, client, storage, tm)

    # test proxy
    if args.test_proxy:
        if not pm or not pm.is_enabled():
            print("Прокси не настроен (proxy.enabled=false)")
        else:
            print(f"Тест прокси: {pm.list_masked()}")
            from .proxy_manager import proxy_string_to_url
            raw=pm.raw_list[0]
            print(f"  строка: {raw}")
            print(f"  URL: {proxy_string_to_url(raw)}")
            ok, msg=client.test_proxy()
            print(f"  результат: {'OK' if ok else 'FAIL'} — {msg}")
        if not args.test:
            return 0

    if args.test:
        inn=args.inn
        print(f"Тест DaData: ИНН {inn} прокси={pm.get_masked() if pm else 'прямое'}")
        try:
            resp, token=client.query_one(inn)
            print(f"  токен: {masked_token(token)}")
            print(f"  http_status: {resp.get('http_status')}")
            if resp.get("http_status")==200:
                sugg=resp.get("data",{}).get("suggestions",[])
                if sugg:
                    d=sugg[0].get("data",{})
                    print(f"  найдено: {d.get('name',{}).get('full_with_opf') or d.get('name',{}).get('full')}")
                    print(f"  руководитель: {d.get('management',{}).get('name')} — {d.get('management',{}).get('post')}")
                    print(f"  статус: {d.get('state',{}).get('status')} ОГРН {d.get('ogrn')} okved {d.get('okved')}")
                else:
                    print("  не найдено (пустой suggestions)")
            else:
                print(f"  ошибка: {resp.get('error') or resp.get('data')}")
            # инкремент для статистики теста
            if resp.get("http_status")==200:
                tm.record_usage(token, inn=inn, http_status=200, note="test")
                print(f"  расход токена: {tm.format_status()}")
            return 0
        except Exception as e:
            print(f"  EXC: {e}")
            import traceback; traceback.print_exc()
            return 1

    # основной прогон
    if not args.input_path and not args.inns:
        parser.print_help()
        print("\nНужно указать --input файл или --inns список")
        return 2

    input_path=args.input_path
    inns=args.inns
    # output resolution
    excel_output=args.excel_output or cfg["output"].get("excel_output")
    # если input — Excel и не задан excel_output — используем из конфига
    # если input — не Excel, output_path — JSONL не используется (всё в storage raw_jsonl)
    try:
        result=worker.run(input_path=input_path, inns=inns, output_path=excel_output, enable_tqdm=not args.no_tqdm)
        print("\nИТОГ:", result)
        print(f"Лог: {cfg['storage']['log_path']} | JSONL: {cfg['storage']['raw_jsonl']} | DB: {cfg['storage']['sqlite_path']}")
        stats=tm.get_stats()
        for s in stats:
            print(f"  {s['token_masked']}: {s['count']}/{s['limit_per_day']} ({s['count']/s['limit_per_day']*100:.1f}%)")
        used,limit=tm.get_total_today()
        print(f"Всего за сегодня: {used}/{limit} ({used/limit*100:.1f}% при limit>0)")
        return 0
    except KeyboardInterrupt:
        print("\nПрервано пользователем")
        return 130
    except Exception as e:
        print(f"ОШИБКА: {e}")
        import traceback; traceback.print_exc()
        return 1

if __name__=="__main__":
    raise SystemExit(main())
