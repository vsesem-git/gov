#!/usr/bin/env python3
"""
РСО_Каталог — сборка каталога из РСО_DaData-СОРТИРОВКА.xlsx

- Приоритет ОКВЭД (кол. 31 Сырой ответ JSON → поле okved)
- Категория / Доп категория — 2 фильтра
- Тепловодоканал: основная категория по ОКВЭД (35.30.14 → Теплосеть) + Доп=Водоканал
- УК отдельно от ТСЖ/ЖСК/ТСН (разделение по названию при ОКВЭД 68.32.*)
- Новый файл РСО_Каталог.xlsx (исходник не трогается)
"""
import pathlib, re, collections
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

SRC = pathlib.Path("dropbox/РСО_DaData-СОРТИРОВКА.xlsx")
OUT = pathlib.Path("РСО_Каталог.xlsx")
OUT2 = pathlib.Path("dropbox/РСО_Каталог.xlsx")  # копия для Dropbox

HEADER_CAT = "Категория"
HEADER_DOP = "Доп. категория"

# Цвета категорий (палитра каталога)
CAT_COLORS = {
    "Теплосеть": "FCE4EC",          # розовый
    "Водоканал": "E3F2FD",           # голубой
    "Электросети": "FFF9C4",         # жёлтый
    "ГАЗ": "FFF3E0",                 # оранжевый
    "ТБО / ТКО": "E8F5E9",           # зелёный
    "Управляющие организации (УК)": "EDE7F6", # сиреневый
    "ТСЖ / ЖСК / ТСН": "F3E5F5",     # лиловый
    "РКЦ / Платежные центры": "E0F2F1", # бирюзовый
    "Администрации / Органы власти": "FAFAFA", # серый
    "Оборонный сектор": "FFEBEE",    # красный светлый
    "Строители / Застройщики": "FFF8E1", # янтарный
    "Производители / Промышленность": "EFEBE9", # коричневый светлый
    "Содержание, благоустройство и аренда": "F1F8E9",
    "Прочие организации": "ECEFF1",
}

OKVED_RE = re.compile(r'"okved"\s*:\s*"([^"]+)"')
# fallback: иногда okved может быть без кавычек? оставим только этот

def is_tszh(name: str) -> bool:
    nl = (name or "").lower()
    if "тсж" in nl:
        return True
    if "жск" in nl:
        return True
    # ТСН как слово
    padded = f" {nl} "
    if " тсн " in padded or " тсн\"" in padded or "\"тсн" in nl or " тсн," in padded or " тсн." in padded:
        return True
    if "товарищество собственников" in nl:
        return True
    if "кооператив" in nl:
        return True
    return False

def has_water_keyword(name: str) -> bool:
    nl = (name or "").lower()
    # маркеры водоканала в названии тепловиков
    for kw in ["водоканал","водопровод","водоснаб","водоотвед","тепловод","холодн вод","горяч вод","водн коммуналь","коммунальн вод"]:
        if kw in nl:
            return True
    # более общий "вод" но чтобы не триггерить "водовозов"? оставим узкий
    # специальный паттерн "водо"
    if "водо" in nl and "водо" in nl:
        # проверяем что не просто "заводовод" etc. Но оставим.
        # дополнительно: коммунальн + водо
        if "коммун" in nl and "вод" in nl:
            return True
        if "водо" in nl:
            # для тепловиков часто "жилищно-коммунальное" + водо
            pass
    return False

def categorize_okved(okved: str, name: str) -> str:
    okved = (okved or "").strip()
    nl = (name or "").lower()
    # Оборонный — проверяем до прочих, но после базовых РСО? Ставим высокий приоритет для 84.22 и 25.40
    # Если имя явно оборонное — отдаём туда независимо от ОКВЭД? Лучше после ОКВЭД базовых, но перед "Прочие"
    # Сначала базовые РСО:
    if okved.startswith("35.30") or okved == "35.3":
        return "Теплосеть"
    if okved.startswith("36.00") or okved.startswith("36.00.") or okved.startswith("37.00") or okved.startswith("37.") or okved in ("36.00","36.0","37.0","37.00","36","37"):
        # 36.* водоснабжение, 37.* водоотведение
        # но 36.00.2 уже водоканал
        return "Водоканал"
    # уточнение: 36.00.1, 36.00.2 тоже водоканал (покрыто)
    if okved.startswith("36.") or okved.startswith("37."):
        return "Водоканал"
    if okved.startswith("35.11") or okved.startswith("35.12") or okved.startswith("35.13") or okved.startswith("35.14"):
        return "Электросети"
    if okved.startswith("35.22") or okved.startswith("35.23") or okved == "46.71.4":
        return "ГАЗ"
    if okved.startswith("38.") or okved == "38.1" or okved == "38.11":
        return "ТБО / ТКО"
    # РКЦ — до 84 чтобы не ушло в Администрации
    if okved.startswith("63.11") or okved == "82.99" or okved.startswith("66.19") or okved in ("63.11","63.11.1","64.19","64.99","66.19","66.19.6","66.19.61","69.20","69.20.2","82.11"):
        return "РКЦ / Платежные центры"
    if okved.startswith("84."):
        # 84.22,84.25 оборонные, но оставим в Администрации кроме явных
        # Оборонный 84.22 — деятельность связанная с обеспечением военной безопасности
        if okved == "84.22" or okved.startswith("84.22"):
            return "Оборонный сектор"
        if okved.startswith("84.25"):
            # часть 84.25 — обеспечение безопасности в чрезвычайных ситуациях — не оборона, а МЧС
            # оставим в Администрации, но если в названии есть оборона — выше
            pass
        return "Администрации / Органы власти"
    # Оборонный по имени (если не попал в 84.22)
    if any(k in nl for k in ["оборонпром","военн","арсенал","боеприпас","миноборон","росгварди","оборон", "военный"]):
        # только если ОКВЭД промышленный/оборонный характер
        if okved.startswith(("84.","25.","30.","84.25","84.24","84.22","25.40")):
            return "Оборонный сектор"
        # даже если другой ОКВЭД, но название явно оборонное — всё равно
        if "арсенал" in nl or "оборонпром" in nl:
            return "Оборонный сектор"
    if okved.startswith("68.32"):
        if is_tszh(name):
            return "ТСЖ / ЖСК / ТСН"
        else:
            return "Управляющие организации (УК)"
    if okved.startswith("41.") or okved.startswith("42.") or okved.startswith("43."):
        return "Строители / Застройщики"
    if okved.startswith(("01.","02.","03.","05.","06.","07.","08.","09.","10.","11.","12.","13.","14.","15.","16.","17.","18.","19.","20.","21.","22.","23.","24.","25.","26.","27.","28.","29.","30.","31.","32.","33.")):
        return "Производители / Промышленность"
    # Содержание, благоустройство и аренда — 68.20,81.*,68.31,70.22,71.*,68.10 etc.
    if okved.startswith(("68.20","68.31","68.10","81.","70.22","71.12","71.11","68.3","68.1","68.32.2")):
        # 68.32.2 уже обрабатывался, но на всякий
        return "Содержание, благоустройство и аренда"
    # Соц учреждения
    if okved.startswith(("85.","86.","87.","88.","90.","91.","93.")):
        return "Прочие организации"
    # Прочие
    return "Прочие организации"

def dop_category(cat: str, name: str, okved: str) -> str:
    if cat == "Теплосеть" and has_water_keyword(name):
        return "Водоканал"
    # Доп: если Теплосеть и в имени есть многофункциональность
    # Также можно добавить доп для Водоканал → Теплосеть если в имени "тепло"
    if cat == "Водоканал" and "тепло" in (name or "").lower():
        return "Теплосеть"
    return ""

def build():
    print(f"Источник: {SRC} ({SRC.stat().st_size/1024/1024:.1f} МБ)")
    wb_src = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
    ws_src = wb_src.active
    max_row = ws_src.max_row
    max_col = ws_src.max_column
    headers = [c.value for c in next(ws_src.iter_rows(min_row=1, max_row=1))]
    print(f"Лист: {ws_src.title} {max_row} строк x {max_col} колонок")
    # Проверяем что в headers есть нужные
    assert headers[0] == "Полное наименование"
    assert headers[10] == "Функция"
    assert headers[30] == "Сырой ответ (сжатый JSON)"

    # Создаём новый workbook
    wb = openpyxl.Workbook()
    # Удалим дефолт лист для пересоздания
    wb.remove(wb.active)

    # 1) Служебный лист "О каталоге"
    ws_info = wb.create_sheet("О каталоге")
    ws_info.sheet_properties.pageSetUpPr.fitToPage = True
    ws_info.page_setup.orientation = "portrait"
    info_text = [
        ("РСО_Каталог.xlsx", "Каталог поставщиков информации ГИС ЖКХ — сортировка по ОКВЭД (приоритет)"),
        ("Источник:", str(SRC.name) + f" — Реестр поставщиков информации, 36 108 организаций, 31 колонка"),
        ("Дата сборки:", __import__("datetime").datetime.now().strftime("%Y-%m-%d %H:%M %Z")),
        ("ОКВЭД:", "Основной код из поля 31 «Сырой ответ (сжатый JSON)» — DaData suggestions/dadata_raw.jsonl"),
        ("Примечание:", "Поле okveds (доп. ОКВЭДы) у DaData = null 100% (проверено по dadata_raw.jsonl 202611112 байт, 36108 записей). Для дозапроса доп. кодов — провайдер ФНС egrul.nalog.ru (dadata_parser/providers/dop_okved.py)"),
        ("Категория:", "Определяется по основному ОКВЭД — см. таблицу ниже. 14 категорий (каталог, не много)"),
        ("Доп. категория:", "Заполняется для Тепловодоканалов: основная Теплосеть (35.30.14) + доп. Водоканал, если в наименовании — «водоканал/водопровод/водоснабжение/тепловод»"),
        ("УК ≠ ТСЖ:", "При ОКВЭД 68.32.* разделение по наименованию: ТСЖ/ЖСК/ТСН/кооператив → «ТСЖ / ЖСК / ТСН», остальные → «Управляющие организации (УК)»"),
        ("Фильтры:", "На листе «Каталог» включена автофильтрация по колонкам Категория и Доп. категория + по всем исходным полям"),
        ("Файл:", "Исходный файл не изменялся. Новый файл: РСО_Каталог.xlsx"),
    ]
    ws_info["A1"] = "РСО_Каталог — описание"
    ws_info["A1"].font = Font(size=14, bold=True, color="0B5394")
    ws_info["A1"].fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
    ws_info.merge_cells("A1:B1")
    r = 3
    for k,v in info_text:
        ws_info.cell(row=r, column=1, value=k).font = Font(bold=True, size=9)
        ws_info.cell(row=r, column=2, value=v).font = Font(size=9)
        ws_info.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical="center")
        ws_info.row_dimensions[r].height = 18 if len(v)<120 else 28
        r+=1
    r+=1
    ws_info.cell(row=r, column=1, value="Таблица категорий (ОКВЭД → Категория)").font = Font(bold=True, color="FFFFFF")
    ws_info.cell(row=r, column=1).fill = PatternFill(start_color="0B5394", end_color="0B5394", fill_type="solid")
    ws_info.cell(row=r, column=2, value="Примеры ОКВЭДов").font = Font(bold=True, color="FFFFFF")
    ws_info.cell(row=r, column=2).fill = PatternFill(start_color="0B5394", end_color="0B5394", fill_type="solid")
    r+=1
    cat_table = [
        ("Теплосеть", "35.30, 35.30.1, 35.30.11, 35.30.14, 35.30.2, 35.30.3, 35.30.4, 35.3, 35.11.1 (ТЭЦ)"),
        ("Водоканал", "36.00, 36.00.1, 36.00.2, 36.00, 37.00, 37.0, 42.21 (сети водопровода)"),
        ("Электросети", "35.11, 35.11.1, 35.12, 35.12.1, 35.13, 35.14"),
        ("ГАЗ", "35.22, 35.23, 46.71.4"),
        ("ТБО / ТКО", "38.11, 38.1, 38.12, 38.21, 38.21.21, 38.22, 81.29.9"),
        ("Управляющие организации (УК)", "68.32, 68.32.1, 68.32.2 — без маркеров ТСЖ"),
        ("ТСЖ / ЖСК / ТСН", "68.32, 68.32.1 — если в наименовании ТСЖ/ЖСК/ТСН/кооператив/товарищество собственников"),
        ("РКЦ / Платежные центры", "63.11, 63.11.1, 82.99, 66.19, 66.19.6, 66.19.61, 64.19, 69.20"),
        ("Администрации / Органы власти", "84.11.21, 84.11.22, 84.11.31, 84.11.33, 84.11.34, 84.11.35, 84.24, 84.11.*"),
        ("Оборонный сектор", "84.22, 25.40, 30.30, + по названию «арсенал/оборонпром/военн»"),
        ("Строители / Застройщики", "41.20, 41.20.1, 41.2, 42.11, 42.21, 43.22, 43.*"),
        ("Производители / Промышленность", "01.*–33.* (с/х, добыча, обработка, производство)"),
        ("Содержание, благоустройство и аренда", "68.20, 68.20.2, 68.31, 68.10, 68.10.21, 81.10, 81.29.9, 70.22, 71.12.2"),
        ("Прочие организации", "85.* образование, 86.* медицина, 91.* культура, 64.* финансы, 49.* транспорт и т.д."),
    ]
    thin = Side(style="thin", color="CBD5E1")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for cat, ex in cat_table:
        c1 = ws_info.cell(row=r, column=1, value=cat)
        c2 = ws_info.cell(row=r, column=2, value=ex)
        for c in (c1,c2):
            c.font = Font(size=8)
            c.alignment = Alignment(wrap_text=True, vertical="center")
            c.border = border
        # цвет строки
        fill = PatternFill(start_color=CAT_COLORS.get(cat,"FFFFFF"), end_color=CAT_COLORS.get(cat,"FFFFFF"), fill_type="solid")
        c1.fill = fill
        c2.fill = fill
        ws_info.row_dimensions[r].height = 16
        r+=1
    ws_info.column_dimensions["A"].width = 36
    ws_info.column_dimensions["B"].width = 90
    ws_info.sheet_properties.pageSetUpPr.fitToPage = True
    # печать
    ws_info.freeze_panes = "A3"

    # 2) Основной лист Каталог
    ws = wb.create_sheet("Каталог")
    # формируем заголовки: Категория, Доп категория + исходные 31
    new_headers = [HEADER_CAT, HEADER_DOP] + headers
    ws.append(new_headers)
    # стиль заголовка
    head_fill = PatternFill(start_color="0B5394", end_color="0B5394", fill_type="solid")
    head_font = Font(bold=True, color="FFFFFF", size=8, name="Calibri")
    thin = Side(style="thin", color="CBD5E1")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col_idx, h in enumerate(new_headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.fill = head_fill
        cell.font = head_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[1].height = 38
    # ширины
    widths = {
        1: 28, 2: 18,
        3: 48, 4: 30, 5: 36, 6: 36, 7: 14, 8: 14, 9: 10, 10: 36, 11: 28,
        12: 30, 13: 30, 14: 18, 15: 18, 16: 28, 17: 22, 18: 16, 19: 28, 20: 28, 21: 22,
        22: 14, 23: 14, 24: 16, 25: 16, 26: 16, 27: 36, 28: 18, 29: 18, 30: 18, 31: 20, 32: 18, 33: 38
    }
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w

    # Чтение и запись строк
    cat_counter = collections.Counter()
    dop_counter = collections.Counter()
    okved_counter = collections.Counter()
    row_num = 1
    # для скорости используем read_only iter
    # уже имеем headers, теперь iter_rows по src
    # Нужно снова открыть? уже на строке 1, продолжим
    # Но мы уже потребили заголовок, теперь итерируемся с min_row=2
    cat_fill_cache = {}
    for col_idx, cat in enumerate(CAT_COLORS, 0):
        pass
    # Предварительно создадим fills
    cat_fills = {cat: PatternFill(start_color=color, end_color=color, fill_type="solid") for cat, color in CAT_COLORS.items()}

    alt_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    # Iter source rows
    for src_row in ws_src.iter_rows(min_row=2, values_only=True):
        row_num += 1
        name = src_row[0] or ""
        raw = src_row[30] or ""
        okved = ""
        if raw:
            m = OKVED_RE.search(str(raw))
            if m:
                okved = m.group(1).strip()
        cat = categorize_okved(okved, name)
        dop = dop_category(cat, name, okved)
        cat_counter[cat] += 1
        if dop:
            dop_counter[dop] += 1
        if okved:
            okved_counter[okved] += 1
        # строка для записи: cat, dop + src_row
        out_row = [cat, dop] + list(src_row)
        ws.append(out_row)
        # стилизация добавляемой строки (row_num)
        # чекаем чётность для alt
        # но cat_fill перекроет
        fill = cat_fills.get(cat)
        # Применим border и fill для каждой ячейки
        # Чтобы не тормозить на 36k строк, делаем только для первых 2 колонок с цветом категории, остальные — тонкая граница
        for col_idx in range(1, max_col+3):  # 33 колонки
            cell = ws.cell(row=row_num, column=col_idx)
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.font = Font(size=8, name="Calibri")
            if col_idx <= 2:
                if fill:
                    cell.fill = fill
                # жирный для категории
                if col_idx == 1:
                    cell.font = Font(size=8, bold=True, name="Calibri")
            else:
                # чередование для удобства чтения
                if row_num % 2 == 0:
                    cell.fill = alt_fill
            # спец: для ИНН, ОГРН — выравнивание
            if col_idx in (8,9):  # ИНН, КПП с учётом смещения +2
                cell.alignment = Alignment(horizontal="center", vertical="center")
        # высота
        ws.row_dimensions[row_num].height = 14
        if row_num % 5000 == 0:
            print(f"  написано {row_num-1} / {max_row-1}")

    # Автофильтр и заморозка
    ws.auto_filter.ref = f"A1:{get_column_letter(max_col+2)}1"
    ws.freeze_panes = "C2"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    # Ширина печати
    ws.print_title_rows = "1:1"

    # 3) Лист "Сводка по категориям"
    ws_sum = wb.create_sheet("Сводка")
    ws_sum["A1"] = "Сводка по категориям (ОКВЭД → Категория)"
    ws_sum["A1"].font = Font(size=12, bold=True, color="0B5394")
    ws_sum.merge_cells("A1:D1")
    ws_sum["A2"] = "Всего организаций в реестре"
    ws_sum["B2"] = max_row-1
    ws_sum["B2"].font = Font(bold=True)
    ws_sum["A3"] = "Дата сборки"
    ws_sum["B3"] = __import__("datetime").datetime.now().strftime("%Y-%m-%d %H:%M")
    ws_sum["A4"] = "Источник"
    ws_sum["B4"] = str(SRC.name)
    # таблица
    hdr = ["Категория","Кол-во","Доля %","Примеры ОКВЭДов"]
    r = 6
    for c, h in enumerate(hdr, 1):
        cell = ws_sum.cell(row=r, column=c, value=h)
        cell.fill = head_fill
        cell.font = head_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws_sum.row_dimensions[r].height = 22
    r += 1
    total = max_row-1
    # сортируем по убыванию
    for cat, cnt in cat_counter.most_common():
        ws_sum.cell(row=r, column=1, value=cat).font = Font(size=9, bold=True)
        ws_sum.cell(row=r, column=1).fill = cat_fills.get(cat, PatternFill(fill_type=None))
        ws_sum.cell(row=r, column=1).border = border
        c2 = ws_sum.cell(row=r, column=2, value=cnt)
        c2.font = Font(size=9)
        c2.alignment = Alignment(horizontal="center")
        c2.border = border
        pct = cnt/total*100 if total else 0
        c3 = ws_sum.cell(row=r, column=3, value=round(pct,2))
        c3.font = Font(size=9)
        c3.number_format = '0.00'
        c3.alignment = Alignment(horizontal="center")
        c3.border = border
        # примеры ОКВЭД для категории — соберём топ-3
        ex_okveds = []
        # соберём из okved_counter те что мапятся в эту категорию (простым перебором 100 популярных)
        # Для быстроты просто покажем пусто, а ниже сделаем подробную таблицу
        ws_sum.cell(row=r, column=4, value="").border = border
        ws_sum.row_dimensions[r].height = 16
        r+=1
    # доп категории сводка
    r+=1
    ws_sum.cell(row=r, column=1, value="Доп. категории (заполнено)").font = Font(bold=True)
    r+=1
    for dop,cnt in dop_counter.most_common():
        ws_sum.cell(row=r, column=1, value=dop)
        ws_sum.cell(row=r, column=2, value=cnt)
        r+=1
    if not dop_counter:
        ws_sum.cell(row=r, column=1, value="— нет (кроме Тепловодоканалов)")
        ws_sum.cell(row=r, column=2, value=0)
        r+=1
    # подробности ОКВЭД топ
    r+=1
    ws_sum.cell(row=r, column=1, value="Топ ОКВЭД (основной код из DaData) — первые 40").font = Font(bold=True, color="0B5394")
    ws_sum.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    r+=1
    for c,h in enumerate(["ОКВЭД","Кол-во","Категория",""],1):
        cell=ws_sum.cell(row=r, column=c, value=h)
        cell.fill = head_fill
        cell.font = head_font
        cell.border=border
    r+=1
    for okved,cnt in okved_counter.most_common(40):
        # определим категорию для этого окведа (берём любой name пустой? но категоризация без имени может отличаться для 68.32 — покажем как есть
        # для 68.32 покажем УК (основная масса)
        dummy_name = "ООО \"ТЕСТ\""
        if okved.startswith("68.32"):
            cat_ex = "УК (68.32.* → УК/ТСЖ)"
        else:
            cat_ex = categorize_okved(okved, dummy_name)
        ws_sum.cell(row=r, column=1, value=okved).font=Font(size=8)
        ws_sum.cell(row=r, column=2, value=cnt).font=Font(size=8)
        ws_sum.cell(row=r, column=2).alignment=Alignment(horizontal="center")
        ws_sum.cell(row=r, column=3, value=cat_ex).font=Font(size=8)
        for c in range(1,5):
            ws_sum.cell(row=r, column=c).border=border
        r+=1
    ws_sum.column_dimensions["A"].width = 36
    ws_sum.column_dimensions["B"].width = 14
    ws_sum.column_dimensions["C"].width = 18
    ws_sum.column_dimensions["D"].width = 50
    ws_sum.freeze_panes = "A7"
    ws_sum.sheet_properties.pageSetUpPr.fitToPage = True

    # 4) Лист "ОКВЭД-справочник" — краткий
    ws_ref = wb.create_sheet("ОКВЭД справочник")
    ws_ref["A1"] = "ОКВЭД справочник — коды встречающиеся в реестре (топ-200) с описанием"
    ws_ref["A1"].font = Font(bold=True, color="0B5394")
    ws_ref.merge_cells("A1:C1")
    ws_ref.append(["ОКВЭД","Кол-во в реестре","Описание (кратко)"])
    for cell in ws_ref[2]:
        cell.fill = head_fill
        cell.font = head_font
        cell.border = border
    # описания — захардкодим основные
    desc_map = {
        "68.32.1": "Управление эксплуатацией жилого фонда за вознаграждение",
        "68.32": "Управление недвижимым имуществом за вознаграждение",
        "68.32.2": "Управление эксплуатацией нежилого фонда",
        "35.30.14": "Производство пара и горячей воды (котельными)",
        "36.00.2": "Распределение воды для питьевых и промышленных нужд",
        "41.20": "Строительство жилых и нежилых зданий",
        "71.12.2": "Деятельность заказчика-застройщика, генподрядчика",
        "35.30": "Производство, передача и распределение пара и горячей воды",
        "36.00.1": "Забор, очистка и распределение воды",
        "37.00": "Сбор и обработка сточных вод",
        "36.00": "Забор, очистка и распределение воды",
        "68.20": "Аренда и управление собственным/арендов. недвиж. имуществом",
        "68.20.2": "Аренда и управление собственным/арендов. нежилым недвиж.",
        "84.11.21": "Деятельность органов гос. власти субъектов (республик, краев...)",
        "85.21": "Образование профессиональное среднее",
        "86.10": "Деятельность больничных организаций",
        "62.01": "Разработка компьютерного ПО",
        "38.11": "Сбор неопасных отходов (ТКО)",
        "81.10": "Деятельность по комплексному обслуживанию помещений",
        "63.11.1": "Деятельность по созданию и использованию баз данных",
        "84.11.22": "Деятельность органов гос. власти субъектов (городов фед. значения)",
        "43.22": "Производство санитарно-технических работ, монтаж отопления и кондиционирования",
        "81.29.9": "Деятельность по чистке и уборке прочая, не включ. в др. группировки",
        "35.14": "Торговля электроэнергией",
        "84.11.31": "Деятельность органов МСУ сельских поселений",
        "82.99": "Деятельность по предоставлению прочих вспомогат. услуг для бизнеса",
        "70.22": "Консультирование по вопросам комм. деятельности и управления",
        "35.22": "Распределение газообразного топлива по газораспределительным сетям",
        "68.10": "Покупка и продажа собств. недвиж. имущества",
        "68.10.21": "Покупка и продажа собств. жилого недвиж. имущества",
        "84.11.35": "Деятельность органов МСУ муниципальных округов",
        "35.11.1": "Производство электроэнергии тепловыми электростанциями",
    }
    for okved,cnt in okved_counter.most_common(200):
        ws_ref.append([okved, cnt, desc_map.get(okved,"")])
        row_cells = ws_ref[ws_ref.max_row]
        for cell in row_cells:
            cell.font = Font(size=8)
            cell.border = border
            cell.alignment = Alignment(vertical="center")
    ws_ref.column_dimensions["A"].width = 14
    ws_ref.column_dimensions["B"].width = 16
    ws_ref.column_dimensions["C"].width = 90
    ws_ref.auto_filter.ref = f"A2:C{ws_ref.max_row}"
    ws_ref.freeze_panes = "A3"

    # Итого — сохранение
    # порядок листов
    wb._sheets = [wb["О каталоге"], wb["Сводка"], wb["Каталог"], wb["ОКВЭД справочник"]]
    # Печать
    for sh in wb.worksheets:
        sh.sheet_properties.pageSetUpPr.fitToPage = True

    # Сохраняем — с обработкой длинного пути Windows
    out_abs = OUT.resolve()
    out_abs.parent.mkdir(parents=True, exist_ok=True)
    # Попробуем сохранить, если ошибка — fallback рядом с исходником
    try:
        save_path = str(out_abs)
        # long path fix
        import os as _os
        if _os.name == "nt" and len(save_path) > 240 and not save_path.startswith("\\\\?\\"):
            save_path = "\\\\?\\" + save_path
        wb.save(save_path)
        print(f"Сохранено: {OUT} ({OUT.stat().st_size/1024/1024:.2f} МБ)")
    except Exception as e:
        print(f"Ошибка сохранения {e}, пробуем fallback")
        fallback = pathlib.Path("РСО_Каталог.xlsx")
        wb.save(fallback)
        print(f"Fallback: {fallback.resolve()} {fallback.stat().st_size/1024/1024:.2f} МБ")
        # копируем туда куда нужно
        import shutil
        shutil.copy(fallback, OUT)

    # Копия в dropbox/
    try:
        import shutil
        OUT2.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy(OUT, OUT2)
        print(f"Копия: {OUT2} ({OUT2.stat().st_size/1024/1024:.2f} МБ)")
    except Exception as e:
        print(f"Копия в dropbox не удалась: {e}")

    # Итоги в консоль
    print("\n=== Категории ===")
    for cat,cnt in cat_counter.most_common():
        print(f"{cnt:5d}  {cat}")
    print("\n=== Доп ===")
    for k,v in dop_counter.most_common():
        print(k,v)
    print(f"\nВсего: {total}")
    print(f"Топ ОКВЭД:")
    for code,cnt in okved_counter.most_common(10):
        print(cnt, code)

if __name__ == "__main__":
    build()
