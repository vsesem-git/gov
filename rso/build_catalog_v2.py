#!/usr/bin/env python3
"""
РСО_Каталог v2 — улучшенные фильтры по названию

- 2 фильтра: 1) Категория по ОКВЭД (из DaData), 2) Категория по фактической деятельности (по названию)
- Основная категория = по названию если есть явный маркер (водоканал, теплосеть и т.д.), иначе по ОКВЭД
- Доп категория = ОКВЭД-категория если отличается от основной (сохранён второй фильтр), + для тепловодоканала Теплосеть+Водоканал
- УК ≠ ТСЖ по-прежнему по названию при 68.32.*
"""
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
# keywords regex for name-based actual activity — strong markers only
# use \b for abbreviations (3 буквы) to avoid "волжская" -> "жск"
NAME_PATTERNS = [
    ("Водоканал", re.compile(r'водоканал|горводоканал|водопровод|водоснабжен|водоотведен', re.I)),
    ("Теплосеть", re.compile(r'теплосет|теплоснаб|теплоэнерго|теплофикац|\bтэц\b|тгк-?\d*|котельная', re.I)),
    ("Электросети", re.compile(r'электросет|облэнерго|горэлектросет|электросбыт|энергосбыт|мрск|россети|фск еэс|\bрэск\b', re.I)),
    ("ГАЗ", re.compile(r'\bгаз\b|облгаз|горгаз|межрегионгаз|газпром|газсеть|газоснабжен', re.I)),
    # ТБО — короткие тко/тбо требуют границ
    ("ТБО / ТКО", re.compile(r'спецавтобаза|спецтранс|спецавтохозяйство|регоператор|региональный оператор.*отход|\bтко\b|\bтбо\b|отход|полигон тбо|эко\s*центр', re.I)),
    ("РКЦ / Платежные центры", re.compile(r'расчетный центр|\bркц\b|\bеркц\b|\bеирц\b|мособлеирц|биллинг|платежный центр|единый расчетный|информационно-расчетн', re.I)),
    ("Администрации / Органы власти", re.compile(r'администрация|мэрия|сельское поселение|городское поселение|муниципальный район|муниципальное образование|совет депутатов', re.I)),
    ("Оборонный сектор", re.compile(r'оборонпром|арсенал|боеприпас|минобороны|военторг|военн', re.I)),
    # ТСЖ/ЖСК — границы!
    ("ТСЖ / ЖСК / ТСН", re.compile(r'\bтсж\b|\bжск\b|\bтсн\b|товарищество собственников|\bкооператив\b', re.I)),
    ("Управляющие организации (УК)", re.compile(r'управляющая компания|управляющая организация|домоуправляющ|жилсервис|жкх-сервис|\bжэу\b|\bжрэу\b', re.I)),
]

def is_tszh(name: str) -> bool:
    return bool(re.search(r'\bтсж\b|\bжск\b|\bтсн\b|товарищество собственников|кооператив', (name or ""), re.I))

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

def categorize_by_name(name: str, func: str = None):
    if func in ("Товарищество собственников жилья","Жилищно-строительный кооператив","Жилищный кооператив","Иной специализированный потребительский кооператив"):
        return "ТСЖ / ЖСК / ТСН"
    if not name: return None
    # специальная обработка тепловодоканал — сразу возвращаем Теплосеть чтобы dop был Водоканал (по старому требованию)
    nl=name.lower()
    if "тепловодоканал" in nl or ("тепло" in nl and "вод" in nl and "водоканал" in nl):
        return "Теплосеть"
    for cat, pat in NAME_PATTERNS:
        if pat.search(name):
            return cat
    return None

def build():
    print(f"SRC {SRC} {SRC.stat().st_size/1024/1024:.1f} MB")
    wb_src = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
    ws_src = wb_src.active
    max_row = ws_src.max_row
    max_col = ws_src.max_column
    headers = [c.value for c in next(ws_src.iter_rows(min_row=1, max_row=1))]
    print(f"sheet {ws_src.title} {max_row} x {max_col}")

    wb = openpyxl.Workbook()
    ws_info = wb.active
    ws_info.title = "О каталоге"
    ws_info["A1"]="РСО_Каталог v2 — описание (улучшенные фильтры по названию)"
    ws_info["A1"].font=Font(size=14, bold=True, color="0B5394")
    ws_info["A1"].fill=PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
    ws_info.merge_cells("A1:B1")
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    thin=Side(style="thin", color="CBD5E1")
    border=Border(left=thin, right=thin, top=thin, bottom=thin)
    head_fill=PatternFill(start_color="0B5394", end_color="0B5394", fill_type="solid")
    head_font=Font(bold=True, color="FFFFFF", size=8, name="Calibri")
    info = [
        ("РСО_Каталог.xlsx","Каталог поставщиков ГИС ЖКХ — ОКВЭД + фактическая деятельность по названию"),
        ("Источник:", f"{SRC.name} — Реестр поставщиков информации, 36 108 организаций, 31 колонка"),
        ("Дата сборки:", now + " — v2 с учётом названия"),
        ("ОКВЭД (фильтр 1):", "Основной код из поля 31 «Сырой ответ JSON» — DaData (okveds=null 100%, проверено dadata_raw.jsonl 202МБ)"),
        ("Факт. деятельность (фильтр 2):", "Если в названии есть явный маркер — «водоканал/водопровод/теплосеть/тэц/электросеть/облэнерго/горгаз/газпром/регоператор/ркц/администрация…» — он уходит в основную категорию по названию, а ОКВЭД остаётся в Доп. категории"),
        ("Пример:", "ОКВЭД 41.20 «Строитель» + название «ВОДОКАНАЛ» → Категория=Водоканал, Доп=Строители. ОКВЭД 35.30.14 «Теплосеть» + «ТЕПЛОВОДОКАНАЛ» → Категория=Теплосеть, Доп=Водоканал (спец-правило)"),
        ("УК ≠ ТСЖ:", "При ОКВЭД 68.32.* ТСЖ/ЖСК/ТСН/кооператив → «ТСЖ / ЖСК / ТСН», остальные → «УК». По названию «ТСЖ/ЖСК» тоже ловится даже без ОКВЭД"),
        ("Фильтры:", "На листе «Каталог» две колонки: Категория (основная, по ОКВЭД+названию) и Доп. категория (второй фильтр) — можно фильтровать каждым отдельно"),
    ]
    r=3
    for k,v in info:
        ws_info.cell(row=r, column=1, value=k).font=Font(bold=True, size=9)
        ws_info.cell(row=r, column=2, value=v).font=Font(size=9)
        ws_info.cell(row=r, column=2).alignment=Alignment(wrap_text=True, vertical="center")
        ws_info.row_dimensions[r].height=26 if len(v)>140 else 16
        r+=1
    r+=1
    ws_info.cell(row=r, column=1, value="Категория").font=Font(bold=True, color="FFFFFF")
    ws_info.cell(row=r, column=1).fill=head_fill
    ws_info.cell(row=r, column=2, value="Маркеры в названии → категория / Примеры ОКВЭДов").font=Font(bold=True, color="FFFFFF")
    ws_info.cell(row=r, column=2).fill=head_fill
    r+=1
    cat_table=[
        ("Теплосеть","теплосеть, теплоснаб, теплоэнерго, тэц, котельная | 35.30, 35.30.14, 35.11.1"),
        ("Водоканал","водоканал, горводоканал, водопровод, водоснабжен, водоотведен | 36.00, 36.00.2, 37.00"),
        ("Электросети","электросеть, облэнерго, электросбыт, энергосбыт, мрск, россети | 35.11, 35.12, 35.14"),
        ("ГАЗ","газ, облгаз, горгаз, межрегионгаз, газпром | 35.22, 35.23"),
        ("ТБО / ТКО","спецавтобаза, спецтранс, регоператор, тко, тбо, отходы | 38.11, 38.1"),
        ("Управляющие организации (УК)","управляющая компания, домоуправляющ, жилсервис | 68.32, 68.32.1"),
        ("ТСЖ / ЖСК / ТСН","\\bтсж\\b, \\bжск\\b, \\bтсн\\b, товарищество собственников | 68.32"),
        ("РКЦ / Платежные центры","расчетный центр, ркц, еркц, еирц, биллинг | 63.11, 63.11.1, 82.99"),
        ("Администрации / Органы власти","администрация, мэрия, сельское поселение, муниципальный район | 84.11.*"),
        ("Оборонный сектор","оборонпром, арсенал, минобороны | 84.22, 25.40"),
        ("Строители / Застройщики","застройщик, строитель, пмк | 41.20, 42.11, 43.*"),
        ("Производители / Промышленность","завод, фабрика, комбинат | 01.*–33.*"),
        ("Содержание, благоустройство и аренда","81.10, 68.20, 68.20.2"),
        ("Прочие организации","прочие"),
    ]
    for cat,ex in cat_table:
        c1=ws_info.cell(row=r, column=1, value=cat)
        c2=ws_info.cell(row=r, column=2, value=ex)
        for c in (c1,c2):
            c.font=Font(size=8); c.alignment=Alignment(wrap_text=True, vertical="center"); c.border=border
        fill=PatternFill(start_color=CAT_COLORS.get(cat,"FFFFFF"), end_color=CAT_COLORS.get(cat,"FFFFFF"), fill_type="solid")
        c1.fill=fill; c2.fill=fill
        ws_info.row_dimensions[r].height=16
        r+=1
    ws_info.column_dimensions["A"].width=36
    ws_info.column_dimensions["B"].width=92
    ws_info.freeze_panes="A3"

    ws_sum = wb.create_sheet("Сводка")
    ws = wb.create_sheet("Каталог")
    new_headers=["Категория","Доп. категория"]+headers
    # добавим ещё скрытую колонку с причиной? нет, оставим коммент в листе
    ws.append(new_headers)
    for col_idx,h in enumerate(new_headers,1):
        cell=ws.cell(row=1, column=col_idx, value=h)
        cell.fill=head_fill; cell.font=head_font; cell.alignment=Alignment(horizontal="center", vertical="center", wrap_text=True); cell.border=border
    ws.row_dimensions[1].height=38
    widths={1:28,2:28, 3:48,4:30,5:36,6:36,7:14,8:14,9:10,10:36,11:28,12:30,13:30,14:18,15:18,16:28,17:22,18:16,19:28,20:28,21:22,22:14,23:14,24:16,25:16,26:16,27:36,28:18,29:18,30:18,31:20,32:18,33:38}
    for col,w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width=w

    cat_counter=collections.Counter()
    dop_counter=collections.Counter()
    okved_counter=collections.Counter()
    name_hit_counter=collections.Counter()
    override_counter=collections.Counter()
    cat_fills={cat: PatternFill(start_color=color, end_color=color, fill_type="solid") for cat,color in CAT_COLORS.items()}
    row_num=1
    # for detailed analysis
    for src_row in ws_src.iter_rows(min_row=2, values_only=True):
        row_num+=1
        name=src_row[0] or ""
        raw=src_row[30] or ""
        okved=""
        if raw:
            m=OKVED_RE.search(str(raw))
            if m: okved=m.group(1).strip()
        cat_okved=categorize_okved(okved, name)
        cat_name=categorize_by_name(name, src_row[10])
        # основная = по названию если есть маркер, иначе по ОКВЭД
        if cat_name:
            cat_main=cat_name
            name_hit_counter[cat_name]+=1
            if cat_name != cat_okved:
                override_counter[f"{cat_okved}→{cat_name}"]+=1
        else:
            cat_main=cat_okved
        # доп: для тепловодоканала особое правило — Теплосеть+Водоканал
        dop=""
        nl=name.lower()
        if cat_main=="Теплосеть" and ("водоканал" in nl or "водопровод" in nl or "водоснаб" in nl or "водоотвед" in nl or "тепловод" in nl):
            dop="Водоканал"
        elif cat_main=="Водоканал" and "тепло" in nl and cat_name=="Водоканал":
            # если водоканал по названию но в имени есть тепло → доп теплосеть
            dop="Теплосеть"
        elif cat_name and cat_name != cat_okved:
            # общий случай: доп = оквэд категория, чтобы сохранить второй фильтр
            dop=cat_okved
        # если cat_main уже из названия, но оквэд тот же — доп пуст
        # для чистоты: если dop == cat_main -> пуст
        if dop == cat_main:
            dop=""
        cat_counter[cat_main]+=1
        if dop: dop_counter[dop]+=1
        if okved: okved_counter[okved]+=1
        out_row=[cat_main, dop]+list(src_row)
        ws.append(out_row)
        cat_cell=ws.cell(row=row_num, column=1)
        dop_cell=ws.cell(row=row_num, column=2)
        fill=cat_fills.get(cat_main)
        if fill:
            cat_cell.fill=fill; cat_cell.font=Font(size=8, bold=True, name="Calibri")
            dop_cell.fill=fill; dop_cell.font=Font(size=8, name="Calibri")
            # если доп заполнен — выделим границу
            if dop:
                dop_cell.fill=PatternFill(start_color="FFFDE7", end_color="FFFDE7", fill_type="solid")
        if row_num % 8000==0:
            print(f" {row_num-1}/{max_row-1}")

    ws.auto_filter.ref=f"A1:{get_column_letter(max_col+2)}1"
    ws.freeze_panes="C2"
    ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.page_setup.orientation="landscape"
    ws.page_setup.paperSize=ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth=1
    ws.print_title_rows="1:1"

    # Сводка
    ws_sum["A1"]="Сводка по категориям v2 (основная = факт. деятельность по названию + ОКВЭД)"
    ws_sum["A1"].font=Font(size=12, bold=True, color="0B5394")
    ws_sum.merge_cells("A1:D1")
    ws_sum["A2"]="Всего организаций"; ws_sum["B2"]=max_row-1; ws_sum["B2"].font=Font(bold=True)
    ws_sum["A3"]="Дата сборки"; ws_sum["B3"]=now+" v2"
    ws_sum["A4"]="Источник"; ws_sum["B4"]=str(SRC.name)
    ws_sum["A5"]=f"Переопределено по названию: {sum(override_counter.values())} из {max_row-1} ({sum(override_counter.values())/(max_row-1)*100:.1f}%) — см. ниже топ переопределений"
    ws_sum["A5"].font=Font(size=8, italic=True, color="64748B")
    hdr=["Категория (основная)","Кол-во","Доля %","Откуда (ОКВЭД/название)"]
    r=7
    for c,h in enumerate(hdr,1):
        cell=ws_sum.cell(row=r, column=c, value=h)
        cell.fill=head_fill; cell.font=head_font; cell.alignment=Alignment(horizontal="center", vertical="center", wrap_text=True); cell.border=border
    r+=1
    total=max_row-1
    for cat,cnt in cat_counter.most_common():
        ws_sum.cell(row=r, column=1, value=cat).font=Font(size=9, bold=True)
        ws_sum.cell(row=r, column=1).fill=cat_fills.get(cat, PatternFill()); ws_sum.cell(row=r, column=1).border=border
        c2=ws_sum.cell(row=r, column=2, value=cnt); c2.font=Font(size=9); c2.alignment=Alignment(horizontal="center"); c2.border=border
        pct=cnt/total*100 if total else 0
        c3=ws_sum.cell(row=r, column=3, value=round(pct,2)); c3.font=Font(size=9); c3.number_format='0.00'; c3.alignment=Alignment(horizontal="center"); c3.border=border
        ws_sum.cell(row=r, column=4, value="").border=border
        ws_sum.row_dimensions[r].height=16
        r+=1
    r+=1
    ws_sum.cell(row=r, column=1, value="Доп. категории (второй фильтр = ОКВЭД если отличается)").font=Font(bold=True)
    r+=1
    for dop,cnt in dop_counter.most_common():
        ws_sum.cell(row=r, column=1, value=dop); ws_sum.cell(row=r, column=2, value=cnt)
        # процент от общего
        ws_sum.cell(row=r, column=3, value=round(cnt/total*100,2)).number_format='0.00'
        r+=1
    if not dop_counter:
        ws_sum.cell(row=r, column=1, value="—"); r+=1
    r+=1
    ws_sum.cell(row=r, column=1, value="Топ переопределений: ОКВЭД → по названию (пример строитель→водоканал)").font=Font(bold=True, color="0B5394")
    ws_sum.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    r+=1
    for c,h in enumerate(["Переопределение","Кол-во","Пример ОКВЭД",""],1):
        cell=ws_sum.cell(row=r, column=c, value=h)
        cell.fill=head_fill; cell.font=head_font; cell.border=border
    r+=1
    for key,cnt in override_counter.most_common(20):
        ws_sum.cell(row=r, column=1, value=key).font=Font(size=8)
        ws_sum.cell(row=r, column=2, value=cnt).font=Font(size=8)
        ws_sum.cell(row=r, column=2).alignment=Alignment(horizontal="center")
        for c in range(1,3): ws_sum.cell(row=r, column=c).border=border
        r+=1
    r+=1
    ws_sum.cell(row=r, column=1, value="Топ ОКВЭД (основной) — 40").font=Font(bold=True, color="0B5394")
    ws_sum.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    r+=1
    for c,h in enumerate(["ОКВЭД","Кол-во","Категория по ОКВЭД","Факт по названию пример"],1):
        cell=ws_sum.cell(row=r, column=c, value=h)
        cell.fill=head_fill; cell.font=head_font; cell.border=border
    r+=1
    for okved,cnt in okved_counter.most_common(40):
        cat_ex="УК (68.32.*)" if okved.startswith("68.32") else categorize_okved(okved,"ООО ТЕСТ")
        ws_sum.cell(row=r, column=1, value=okved).font=Font(size=8)
        ws_sum.cell(row=r, column=2, value=cnt).font=Font(size=8)
        ws_sum.cell(row=r, column=2).alignment=Alignment(horizontal="center")
        ws_sum.cell(row=r, column=3, value=cat_ex).font=Font(size=8)
        for c in range(1,4): ws_sum.cell(row=r, column=c).border=border
        r+=1
    ws_sum.column_dimensions["A"].width=38
    ws_sum.column_dimensions["B"].width=14
    ws_sum.column_dimensions["C"].width=18
    ws_sum.column_dimensions["D"].width=52
    ws_sum.freeze_panes="A8"

    ws_ref=wb.create_sheet("ОКВЭД справочник")
    ws_ref["A1"]="ОКВЭД справочник — коды из реестра (топ-200)"; ws_ref["A1"].font=Font(bold=True, color="0B5394")
    ws_ref.merge_cells("A1:C1")
    ws_ref.append(["ОКВЭД","Кол-во","Описание"])
    for cell in ws_ref[2]: cell.fill=head_fill; cell.font=head_font; cell.border=border
    desc_map={"68.32.1":"Управление жилфондом","68.32":"Управление недвижимостью","35.30.14":"Производство пара/горячей воды","36.00.2":"Распределение воды","41.20":"Строительство","71.12.2":"Заказчик-застройщик","35.30":"Производство пара","36.00.1":"Забор воды","37.00":"Сточные воды","68.20":"Аренда недвижимости","38.11":"Сбор отходов ТКО","63.11.1":"РКЦ БД","84.11.21":"Органы гос власти","35.22":"Распределение газа","35.11.1":"ТЭЦ электроэнергия"}
    for okved,cnt in okved_counter.most_common(200):
        ws_ref.append([okved, cnt, desc_map.get(okved,"")])
        for cell in ws_ref[ws_ref.max_row]: cell.font=Font(size=8); cell.border=border
    ws_ref.column_dimensions["A"].width=14; ws_ref.column_dimensions["B"].width=14; ws_ref.column_dimensions["C"].width=90
    ws_ref.auto_filter.ref=f"A2:C{ws_ref.max_row}"; ws_ref.freeze_panes="A3"

    wb._sheets=[wb["О каталоге"], wb["Сводка"], wb["Каталог"], wb["ОКВЭД справочник"]]
    print("Сохранение...")
    wb.save(OUT)
    print(f"OK {OUT} {OUT.stat().st_size/1024/1024:.2f} MB")
    import shutil
    try: shutil.copy(OUT, OUT2); print(f"copy {OUT2} {OUT2.stat().st_size/1024/1024:.2f} MB")
    except Exception as e: print(e)
    print("\nКатегории основные:")
    for cat,cnt in cat_counter.most_common(): print(f"{cnt:5d} {cat}")
    print("\nДоп:", dict(dop_counter))
    print(f"\nПереопределений по названию: {sum(override_counter.values())}")
    for k,v in override_counter.most_common(15): print(f" {v:4d} {k}")
    print(f"Всего {total}")

if __name__=="__main__":
    build()
