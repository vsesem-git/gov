#!/usr/bin/env python3
import pathlib, re, collections, datetime
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SRC = pathlib.Path("dropbox/РСО_DaData-СОРТИРОВКА.xlsx")
OUT = pathlib.Path("РСО_Каталог.xlsx")
OUT2 = pathlib.Path("dropbox/РСО_Каталог.xlsx")

OKVED_RE = re.compile(r'"okved"\s*:\s*"([^"]+)"')
CAT_COLORS = {
    "Теплосеть": "FCE4EC",
    "Водоканал": "E3F2FD",
    "Электросети": "FFF9C4",
    "ГАЗ": "FFF3E0",
    "ТБО / ТКО": "E8F5E9",
    "Управляющие организации (УК)": "EDE7F6",
    "ТСЖ / ЖСК / ТСН": "F3E5F5",
    "РКЦ / Платежные центры": "E0F2F1",
    "Администрации / Органы власти": "FAFAFA",
    "Оборонный сектор": "FFEBEE",
    "Строители / Застройщики": "FFF8E1",
    "Производители / Промышленность": "EFEBE9",
    "Содержание, благоустройство и аренда": "F1F8E9",
    "Прочие организации": "ECEFF1",
}

def is_tszh(name: str) -> bool:
    nl = (name or "").lower()
    if "тсж" in nl: return True
    if "жск" in nl: return True
    padded = f" {nl} "
    if " тсн " in padded or " тсн\"" in padded or "\"тсн" in nl: return True
    if "товарищество собственников" in nl: return True
    if "кооператив" in nl: return True
    return False

def has_water_keyword(name: str) -> bool:
    nl = (name or "").lower()
    for kw in ["водоканал","водопровод","водоснаб","водоотвед","тепловод"]:
        if kw in nl: return True
    return False

def categorize_okved(okved: str, name: str) -> str:
    okved=(okved or "").strip()
    nl=(name or "").lower()
    if okved.startswith("35.30") or okved=="35.3": return "Теплосеть"
    if okved.startswith("36.") or okved.startswith("37.") or okved in ("36.0","37.0","36.00","37.00"): return "Водоканал"
    if okved.startswith("35.11") or okved.startswith("35.12") or okved.startswith("35.13") or okved.startswith("35.14"): return "Электросети"
    if okved.startswith("35.22") or okved.startswith("35.23") or okved=="46.71.4": return "ГАЗ"
    if okved.startswith("38.") or okved=="38.1": return "ТБО / ТКО"
    if okved.startswith("63.11") or okved=="82.99" or okved.startswith("66.19") or okved in ("63.11","63.11.1","64.19","64.99","66.19","66.19.6","66.19.61","69.20","69.20.2","82.11"): return "РКЦ / Платежные центры"
    if okved.startswith("84."):
        if okved=="84.22" or okved.startswith("84.22"): return "Оборонный сектор"
        return "Администрации / Органы власти"
    if any(k in nl for k in ["оборонпром","военн","арсенал","боеприпас"]) and okved.startswith(("84.","25.","30.")): return "Оборонный сектор"
    if "арсенал" in nl or "оборонпром" in nl: return "Оборонный сектор"
    if okved.startswith("68.32"):
        return "ТСЖ / ЖСК / ТСН" if is_tszh(name) else "Управляющие организации (УК)"
    if okved.startswith("41.") or okved.startswith("42.") or okved.startswith("43."): return "Строители / Застройщики"
    if okved.startswith(("01.","02.","03.","05.","06.","07.","08.","09.","10.","11.","12.","13.","14.","15.","16.","17.","18.","19.","20.","21.","22.","23.","24.","25.","26.","27.","28.","29.","30.","31.","32.","33.")): return "Производители / Промышленность"
    if okved.startswith(("68.20","68.31","68.10","81.","70.22","71.12","71.11","68.3","68.1")): return "Содержание, благоустройство и аренда"
    return "Прочие организации"

def dop_category(cat, name, okved):
    if cat=="Теплосеть" and has_water_keyword(name): return "Водоканал"
    if cat=="Водоканал" and "тепло" in (name or "").lower(): return "Теплосеть"
    return ""

def build():
    print(f"SRC {SRC} {SRC.stat().st_size/1024/1024:.1f} MB")
    wb_src = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
    ws_src = wb_src.active
    max_row = ws_src.max_row
    max_col = ws_src.max_column
    headers = [c.value for c in next(ws_src.iter_rows(min_row=1, max_row=1))]
    print(f"sheet {ws_src.title} {max_row} x {max_col}")

    wb = openpyxl.Workbook()
    # --- О каталоге ---
    ws_info = wb.active
    ws_info.title = "О каталоге"
    ws_info["A1"]="РСО_Каталог — описание"
    ws_info["A1"].font=Font(size=14, bold=True, color="0B5394")
    ws_info["A1"].fill=PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
    ws_info.merge_cells("A1:B1")
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    info = [
        ("РСО_Каталог.xlsx","Каталог поставщиков ГИС ЖКХ — сортировка по ОКВЭД (приоритет)"),
        ("Источник:", f"{SRC.name} — Реестр поставщиков информации, 36 108 организаций, 31 колонка"),
        ("Дата сборки:", now),
        ("ОКВЭД:", "Основной код из поля 31 «Сырой ответ (сжатый JSON)» — DaData"),
        ("Примечание:", "okveds (доп. ОКВЭДы) у DaData = null 100% (проверено dadata_raw.jsonl 202611112 байт). Для дозагрузки — провайдер ФНС egrul.nalog.ru (dadata_parser/providers/dop_okved.py)"),
        ("Категория:", "По основному ОКВЭД — 14 категорий (каталог, не много)"),
        ("Доп. категория:", "Для Тепловодоканалов: основная Теплосеть (35.30.14) + доп. Водоканал, если в названии есть водоканал/водопровод/водоснабжение/тепловод"),
        ("УК ≠ ТСЖ:", "При ОКВЭД 68.32.*: ТСЖ/ЖСК/ТСН/кооператив → «ТСЖ / ЖСК / ТСН», остальные → «УК»"),
        ("Фильтры:", "Автофильтр по Категория и Доп. категория (+ все поля) на листе «Каталог»"),
    ]
    r=3
    for k,v in info:
        ws_info.cell(row=r, column=1, value=k).font=Font(bold=True, size=9)
        ws_info.cell(row=r, column=2, value=v).font=Font(size=9)
        ws_info.cell(row=r, column=2).alignment=Alignment(wrap_text=True, vertical="center")
        ws_info.row_dimensions[r].height=22 if len(v)>120 else 16
        r+=1
    r+=1
    ws_info.cell(row=r, column=1, value="Категория").font=Font(bold=True, color="FFFFFF")
    ws_info.cell(row=r, column=1).fill=PatternFill(start_color="0B5394", end_color="0B5394", fill_type="solid")
    ws_info.cell(row=r, column=2, value="Примеры ОКВЭДов").font=Font(bold=True, color="FFFFFF")
    ws_info.cell(row=r, column=2).fill=PatternFill(start_color="0B5394", end_color="0B5394", fill_type="solid")
    r+=1
    cat_table=[
        ("Теплосеть","35.30, 35.30.1, 35.30.14, 35.30.2, 35.30.3, 35.30.4, 35.3, 35.11.1 (ТЭЦ)"),
        ("Водоканал","36.00, 36.00.1, 36.00.2, 37.00, 42.21"),
        ("Электросети","35.11, 35.11.1, 35.12, 35.12.1, 35.13, 35.14"),
        ("ГАЗ","35.22, 35.23, 46.71.4"),
        ("ТБО / ТКО","38.11, 38.1, 38.12, 38.21, 38.21.21, 38.22"),
        ("Управляющие организации (УК)","68.32, 68.32.1 — без маркеров ТСЖ"),
        ("ТСЖ / ЖСК / ТСН","68.32 — ТСЖ/ЖСК/ТСН/кооператив в названии"),
        ("РКЦ / Платежные центры","63.11, 63.11.1, 82.99, 66.19, 69.20"),
        ("Администрации / Органы власти","84.11.21, 84.11.22, 84.11.31, 84.11.33, 84.11.34, 84.24"),
        ("Оборонный сектор","84.22, 25.40, + по названию арсенал/оборонпром"),
        ("Строители / Застройщики","41.20, 41.2, 42.11, 42.21, 43.*"),
        ("Производители / Промышленность","01.*–33.*"),
        ("Содержание, благоустройство и аренда","68.20, 68.20.2, 68.31, 68.10, 81.10, 81.29.9"),
        ("Прочие организации","85.* образование, 86.* медицина, 64.* финансы, 49.* транспорт и др."),
    ]
    thin=Side(style="thin", color="CBD5E1")
    border=Border(left=thin, right=thin, top=thin, bottom=thin)
    for cat,ex in cat_table:
        c1=ws_info.cell(row=r, column=1, value=cat)
        c2=ws_info.cell(row=r, column=2, value=ex)
        for c in (c1,c2):
            c.font=Font(size=8)
            c.alignment=Alignment(wrap_text=True, vertical="center")
            c.border=border
        fill=PatternFill(start_color=CAT_COLORS.get(cat,"FFFFFF"), end_color=CAT_COLORS.get(cat,"FFFFFF"), fill_type="solid")
        c1.fill=fill; c2.fill=fill
        ws_info.row_dimensions[r].height=16
        r+=1
    ws_info.column_dimensions["A"].width=36
    ws_info.column_dimensions["B"].width=90
    ws_info.freeze_panes="A3"

    # --- Сводка (создадим позже после подсчёта) ---
    ws_sum = wb.create_sheet("Сводка")

    # --- Каталог ---
    ws = wb.create_sheet("Каталог")
    new_headers=["Категория","Доп. категория"]+headers
    ws.append(new_headers)
    head_fill=PatternFill(start_color="0B5394", end_color="0B5394", fill_type="solid")
    head_font=Font(bold=True, color="FFFFFF", size=8, name="Calibri")
    for col_idx,h in enumerate(new_headers,1):
        cell=ws.cell(row=1, column=col_idx, value=h)
        cell.fill=head_fill
        cell.font=head_font
        cell.alignment=Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border=border
    ws.row_dimensions[1].height=38
    widths={1:28,2:18, 3:48,4:30,5:36,6:36,7:14,8:14,9:10,10:36,11:28,12:30,13:30,14:18,15:18,16:28,17:22,18:16,19:28,20:28,21:22,22:14,23:14,24:16,25:16,26:16,27:36,28:18,29:18,30:18,31:20,32:18,33:38}
    for col,w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width=w

    cat_counter=collections.Counter()
    dop_counter=collections.Counter()
    okved_counter=collections.Counter()
    cat_fills={cat: PatternFill(start_color=color, end_color=color, fill_type="solid") for cat,color in CAT_COLORS.items()}

    row_num=1
    batch=[]
    # Для скорости — не делаем border для каждой ячейки, только header и категории
    for src_row in ws_src.iter_rows(min_row=2, values_only=True):
        row_num+=1
        name=src_row[0] or ""
        raw=src_row[30] or ""
        okved=""
        if raw:
            m=OKVED_RE.search(str(raw))
            if m: okved=m.group(1).strip()
        cat=categorize_okved(okved, name)
        dop=dop_category(cat, name, okved)
        cat_counter[cat]+=1
        if dop: dop_counter[dop]+=1
        if okved: okved_counter[okved]+=1
        out_row=[cat,dop]+list(src_row)
        ws.append(out_row)
        # Стилизуем только первые 2 ячейки (категории) — остальные оставляем без стилей для скорости
        cat_cell=ws.cell(row=row_num, column=1)
        dop_cell=ws.cell(row=row_num, column=2)
        fill=cat_fills.get(cat)
        if fill:
            cat_cell.fill=fill
            cat_cell.font=Font(size=8, bold=True, name="Calibri")
            dop_cell.fill=fill
            dop_cell.font=Font(size=8, name="Calibri")
        # высота
        # не ставим для каждой строки border — быстрее
        if row_num % 8000==0:
            print(f" {row_num-1}/{max_row-1}  {cat_counter.most_common(3)}")

    ws.auto_filter.ref=f"A1:{get_column_letter(max_col+2)}1"
    ws.freeze_panes="C2"
    ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.page_setup.orientation="landscape"
    ws.page_setup.paperSize=ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth=1
    ws.page_setup.fitToHeight=0
    ws.print_title_rows="1:1"

    # --- Сводка наполнение ---
    ws_sum["A1"]="Сводка по категориям (ОКВЭД → Категория)"
    ws_sum["A1"].font=Font(size=12, bold=True, color="0B5394")
    ws_sum.merge_cells("A1:D1")
    ws_sum["A2"]="Всего организаций"
    ws_sum["B2"]=max_row-1
    ws_sum["B2"].font=Font(bold=True)
    ws_sum["A3"]="Дата сборки"
    ws_sum["B3"]=now
    ws_sum["A4"]="Источник"
    ws_sum["B4"]=str(SRC.name)
    hdr=["Категория","Кол-во","Доля %","Примеры ОКВЭДов"]
    r=6
    for c,h in enumerate(hdr,1):
        cell=ws_sum.cell(row=r, column=c, value=h)
        cell.fill=head_fill; cell.font=head_font
        cell.alignment=Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border=border
    r+=1
    total=max_row-1
    for cat,cnt in cat_counter.most_common():
        ws_sum.cell(row=r, column=1, value=cat).font=Font(size=9, bold=True)
        ws_sum.cell(row=r, column=1).fill=cat_fills.get(cat, PatternFill())
        ws_sum.cell(row=r, column=1).border=border
        c2=ws_sum.cell(row=r, column=2, value=cnt)
        c2.font=Font(size=9); c2.alignment=Alignment(horizontal="center"); c2.border=border
        pct=cnt/total*100 if total else 0
        c3=ws_sum.cell(row=r, column=3, value=round(pct,2))
        c3.font=Font(size=9); c3.number_format='0.00'; c3.alignment=Alignment(horizontal="center"); c3.border=border
        ws_sum.cell(row=r, column=4, value="").border=border
        ws_sum.row_dimensions[r].height=16
        r+=1
    r+=1
    ws_sum.cell(row=r, column=1, value="Доп. категории (заполнено)").font=Font(bold=True)
    r+=1
    for dop,cnt in dop_counter.most_common():
        ws_sum.cell(row=r, column=1, value=dop)
        ws_sum.cell(row=r, column=2, value=cnt)
        r+=1
    if not dop_counter:
        ws_sum.cell(row=r, column=1, value="— только Тепловодоканалы (Теплосеть+Водоканал)")
        ws_sum.cell(row=r, column=2, value=0)
        r+=1
    else:
        # покажем пример тепловодоканалов
        pass
    r+=1
    ws_sum.cell(row=r, column=1, value="Топ ОКВЭД (основной) — 40").font=Font(bold=True, color="0B5394")
    ws_sum.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    r+=1
    for c,h in enumerate(["ОКВЭД","Кол-во","Категория"],1):
        cell=ws_sum.cell(row=r, column=c, value=h)
        cell.fill=head_fill; cell.font=head_font; cell.border=border
    r+=1
    for okved,cnt in okved_counter.most_common(40):
        cat_ex="УК (68.32.*)" if okved.startswith("68.32") else categorize_okved(okved,"ООО ТЕСТ")
        ws_sum.cell(row=r, column=1, value=okved).font=Font(size=8)
        ws_sum.cell(row=r, column=2, value=cnt).font=Font(size=8)
        ws_sum.cell(row=r, column=2).alignment=Alignment(horizontal="center")
        ws_sum.cell(row=r, column=3, value=cat_ex).font=Font(size=8)
        for c in range(1,4):
            ws_sum.cell(row=r, column=c).border=border
        r+=1
    ws_sum.column_dimensions["A"].width=38
    ws_sum.column_dimensions["B"].width=14
    ws_sum.column_dimensions["C"].width=18
    ws_sum.column_dimensions["D"].width=50
    ws_sum.freeze_panes="A7"

    # --- ОКВЭД справочник ---
    ws_ref=wb.create_sheet("ОКВЭД справочник")
    ws_ref["A1"]="ОКВЭД справочник — коды из реестра (топ-200)"
    ws_ref["A1"].font=Font(bold=True, color="0B5394")
    ws_ref.merge_cells("A1:C1")
    ws_ref.append(["ОКВЭД","Кол-во","Описание"])
    for cell in ws_ref[2]:
        cell.fill=head_fill; cell.font=head_font; cell.border=border
    desc_map={
        "68.32.1":"Управление эксплуатацией жилого фонда",
        "68.32":"Управление недвижимым имуществом",
        "35.30.14":"Производство пара и горячей воды (котельными)",
        "36.00.2":"Распределение воды для питьевых нужд",
        "41.20":"Строительство жилых и нежилых зданий",
        "71.12.2":"Деятельность заказчика-застройщика",
        "35.30":"Производство/передача пара и горячей воды",
        "36.00.1":"Забор, очистка и распределение воды",
        "37.00":"Сбор и обработка сточных вод",
        "68.20":"Аренда и управление недвижимостью",
        "38.11":"Сбор неопасных отходов (ТКО)",
        "63.11.1":"Деятельность по созданию и использованию БД/РКЦ",
        "84.11.21":"Деятельность органов гос. власти субъектов",
        "35.22":"Распределение газа",
        "35.11.1":"Производство электроэнергии (ТЭЦ)",
    }
    for okved,cnt in okved_counter.most_common(200):
        ws_ref.append([okved, cnt, desc_map.get(okved,"")])
        for cell in ws_ref[ws_ref.max_row]:
            cell.font=Font(size=8); cell.border=border
    ws_ref.column_dimensions["A"].width=14
    ws_ref.column_dimensions["B"].width=14
    ws_ref.column_dimensions["C"].width=90
    ws_ref.auto_filter.ref=f"A2:C{ws_ref.max_row}"
    ws_ref.freeze_panes="A3"

    wb._sheets=[wb["О каталоге"], wb["Сводка"], wb["Каталог"], wb["ОКВЭД справочник"]]
    print("Сохранение...")
    wb.save(OUT)
    print(f"OK {OUT} {OUT.stat().st_size/1024/1024:.2f} MB")
    import shutil
    try:
        shutil.copy(OUT, OUT2)
        print(f"copy {OUT2} {OUT2.stat().st_size/1024/1024:.2f} MB")
    except Exception as e: print(e)
    print("\nКатегории:")
    for cat,cnt in cat_counter.most_common():
        print(f"{cnt:5d} {cat}")
    print("Доп:", dict(dop_counter))
    print(f"Всего {total}")

if __name__=="__main__":
    build()
