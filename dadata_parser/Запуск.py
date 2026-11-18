#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Запуск.py — один двойной клик для Windows
• Проверяет Python, ставит зависимости (openpyxl, requests[socks], pyyaml, tqdm)
• Настраивает токены и прокси ВНУТРИ GUI — без cmd вообще
• Запускает GUI dadata_parser

Просто: двойной клик по файлу → окно открылось → вставляете токены/прокси → Запустить
Если .py не ассоциирован с Python: используйте Запустить.bat
"""
import sys, os, pathlib, subprocess, importlib.util

# Фикс кодировки на Windows
if sys.platform == "win32":
    try:
        os.system("chcp 65001 > nul")
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except: pass

ROOT = pathlib.Path(__file__).resolve().parent
REQ = ROOT / "requirements.txt"
CFG_EXAMPLE = ROOT / "config.yaml.example"
CFG = ROOT / "config.yaml"

REQUIRED = ["openpyxl", "requests", "socks", "yaml", "tqdm"]

def has_pkg(name):
    return importlib.util.find_spec(name) is not None

def ensure_deps():
    missing = [p for p in ["openpyxl", "requests", "yaml", "tqdm"] if not has_pkg(p)]
    # socks — это pysocks, проверяется как 'socks'
    has_socks = has_pkg("socks")
    if not has_socks:
        missing.append("requests[socks]")
    if not missing:
        return True
    print(f"[*] Ставлю зависимости: {', '.join(missing)} ...")
    print(f"    pip install -r {REQ}")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "pip"], stdout=subprocess.DEVNULL)
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", str(REQ)])
        print("[OK] Зависимости установлены")
        return True
    except subprocess.CalledProcessError as e:
        print(f"[ERR] pip install не удался: {e}")
        # пробуем --user
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "--user", "-r", str(REQ)])
            print("[OK] Зависимости установлены (--user)")
            return True
        except Exception as e2:
            print(f"[ERR] {e2}")
            return False

def ensure_config():
    if CFG.exists():
        return
    if CFG_EXAMPLE.exists():
        import shutil
        shutil.copy(str(CFG_EXAMPLE), str(CFG))
        print(f"[*] Создан {CFG} из примера — токен и прокси можно вставить в GUI")
    else:
        CFG.write_text("# см. config.yaml.example\n", encoding="utf-8")

def launch_gui():
    # Добавляем ROOT в sys.path чтобы импортировать пакет
    if str(ROOT) not in sys.path:
        sys.path.insert(0, str(ROOT))
    ensure_deps()
    ensure_config()
    try:
        from dadata_parser.gui.app import main
        print("[*] Запускаю GUI ...")
        main()
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"\n[ERR] Не удалось запустить GUI: {e}")
        input("\nНажмите Enter для выхода...")
        sys.exit(1)

if __name__ == "__main__":
    print("="*60)
    print(" DaData Parser — запуск GUI")
    print(f" Python {sys.version.split()[0]} | {ROOT}")
    print("="*60)
    launch_gui()
