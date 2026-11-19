#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Пересобрать Excel из уже спарсенного JSONL — без повторных запросов к DaData.
Если при парсинге файл не создался (длинный путь, права, формат реестра), запустите это.

Использование (двойной клик или в консоли):
  python переcобрать_excel.py

Или указать пути:
  python пересобрать_excel.py --input "C:/.../Реестр поставщиков информации от 2026-11-10.xlsx" --jsonl "data/dadata_raw.jsonl" --output "C:/.../Реестр_обогащен.xlsx"

Ищет JSONL автоматически в dadata_parser/data/ и рядом.
"""
import pathlib, sys, json, argparse

ROOT = pathlib.Path(__file__).resolve().parent

def find_jsonl():
    cands=[
        ROOT / "data" / "dadata_raw.jsonl",
        ROOT / "dadata_parser" / "data" / "dadata_raw.jsonl",
        pathlib.Path("data/dadata_raw.jsonl"),
        pathlib.Path("dadata_raw.jsonl"),
    ]
    for p in cands:
        if p.exists() and p.stat().st_size>0:
            return p
    return None

def main():
    p=argparse.ArgumentParser(description="Пересобрать Excel из JSONL без повторного парсинга")
    p.add_argument("--input", default=None, help="Входной реестр (xlsx с ИНН)")
    p.add_argument("--jsonl", default=None, help="Путь к dadata_raw.jsonl")
    p.add_argument("--output", default=None, help="Куда сохранить обогащённый xlsx")
    args=p.parse_args()

    # поиск input
    if not args.input:
        # ищем рядом Реестр*.xlsx
        cands=list(ROOT.glob("*.xlsx")) + list(ROOT.glob("*.xls")) + list((ROOT/"data").glob("*.xlsx"))
        # приоритет — последний Реестр поставщиков
        if cands:
            args.input=str(cands[0])
        else:
            print("Укажите --input путь к исходному реестру (xlsx с колонкой ИНН)")
            print("Например: python пересобрать_excel.py --input \"C:/Downloads/Реестр поставщиков...xlsx\"")
            return 1

    jsonl=pathlib.Path(args.jsonl) if args.jsonl else find_jsonl()
    if not jsonl or not jsonl.exists():
        print(f"JSONL не найден. Укажите --jsonl путь к data/dadata_raw.jsonl")
        print(f"Искал: {find_jsonl()}")
        return 1

    # загрузка результатов из JSONL
    print(f"[1/3] Читаю JSONL: {jsonl} ({jsonl.stat().st_size/1024/1024:.2f} МБ)")
    results={}
    with open(jsonl, encoding="utf-8") as f:
        for line in f:
            line=line.strip()
            if not line: continue
            try:
                obj=json.loads(line)
                inn=str(obj.get("inn") or "").strip()
                resp=obj.get("response") or {}
                if inn: results[inn]=resp
            except: continue
    print(f"  Загружено {len(results)} записей из JSONL")

    output=args.output
    if not output:
        inp=pathlib.Path(args.input)
        output=str(inp.parent / f"{inp.stem}_обогащен_DaData.xlsx")
    print(f"[2/3] Собираю Excel: {args.input} -> {output}")
    sys.path.insert(0, str(ROOT))
    try:
        from dadata_parser.worker import enrich_excel
    except ImportError:
        sys.path.insert(0, str(ROOT/"dadata_parser"))
        from worker import enrich_excel

    try:
        out=enrich_excel(args.input, output, results)
        print(f"[3/3] ✅ Готово: {out}")
        print(f"  Размер: {pathlib.Path(out).stat().st_size/1024/1024:.2f} МБ")
        print(f"  Откройте в Excel")
        # также CSV fallback
        from dadata_parser.worker import save_fallback_csv
        csv_path=pathlib.Path(out).with_suffix(".csv")
        save_fallback_csv(results, csv_path)
        print(f"  CSV копия: {csv_path} ({csv_path.stat().st_size/1024:.1f} КБ)")
        return 0
    except Exception as e:
        import traceback
        print(f"❌ Ошибка: {e}")
        traceback.print_exc()
        print("\nПопробуйте сохранить в короткую папку без русских букв: C:/temp/out.xlsx")
        return 1

if __name__=="__main__":
    raise SystemExit(main())
