#!/usr/bin/env python3
"""
Обогащает РСО_DaData.xlsx механизмом разбивки по назначению.
Источник истины — ОКВЭД-2 из DaData (поле okved), а не название и не поле ГИС ЖКХ 'Функция' (у всех одно значение 'Ресурсоснабжающая организация').

Категории:
  TEPLO — 35.30* (теплоснабжение, котельные, теплосети)
  VODA  — 36.* / 37.* (водоснабжение, водоотведение, Водоканал)
  GAZ   — 35.22 / 35.23 / 49.50.21 (газоснабжение, облгаз/горгаз)
  ELEC  — 35.11* / 35.12* / 35.13 / 35.14 (электроснабжение, сети, сбыт)
  TKO   — 38.11 / 38.12 (регоператор ТКО)
  NON_RSO — всё остальное (ошибочная роль в ГИС ЖКХ: УК, школы, санатории, заводы, фермы и т.д.)

Добавляет колонки:
  AD-AH — Категория РСО, Код категории, ОКВЭД (DaData), Расшифровка ОКВЭД, Раздел ОКВЭД, Статус ЕГРЮЛ, Признак мусорной записи, Комментарий
+ условное форматирование
"""
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, FormulaRule
import json

SRC_XLSX = "/home/user/gov/rso/РСО_DaData.xlsx"
DST_XLSX = "/home/user/gov/rso/РСО_DaData_категоризировано.xlsx"
JSONL = "/home/user/gov/dropbox/РСО.jsonl"

# --- ОКВЭД справочник (сокращённо для отчёта) ---
OKVED_HUMAN = {
 "35.30":"Производство, передача и распределение пара и гор.воды (группа)",
 "35.30.1":"Производство пара и горячей воды (в т.ч. котельными)",
 "35.30.11":"Производство пара/гор.воды ТЭЦ",
 "35.30.12":"Производство пара/гор.воды котельными",
 "35.30.14":"Обеспечение работоспособности котельных",
 "35.30.2":"Передача пара и горячей воды",
 "35.30.3":"Распределение пара и горячей воды",
 "35.30.4":"Обеспечение работоспособности тепловых сетей",
 "35.30.5":"Обеспечение работоспособности тепловых сетей (прочее)",
 "35.30.6":"Торговля паром и горячей водой",
 "35.30.15":"Обеспечение работоспособности котельных (устар.)",
 "35.3":"Устар. = 35.30",
 "36.00":"Забор, очистка и распределение воды (группа)",
 "36.00.1":"Забор и очистка воды",
 "36.00.2":"Распределение воды",
 "36.00.3":"Распределение воды (прочее)",
 "36.0":"Устар. = 36.00",
 "37.00":"Сбор и обработка сточных вод (водоотведение/канализация)",
 "37.0":"Устар. = 37.00",
 "35.22":"Распределение газообразного топлива (облгаз/горгаз)",
 "35.23":"Торговля газом по сетям",
 "49.50.21":"Транспортирование газа по трубопроводам",
 "49.50.2":"Транспортирование по трубопроводам (проч.)",
 "35.11":"Производство электроэнергии",
 "35.11.1":"Производство э/э ТЭС (ТЭЦ/ГРЭС)",
 "35.12":"Передача электроэнергии",
 "35.12.1":"Передача э/э (электросети)",
 "35.12.2":"Распределение э/э (альт.)",
 "35.13":"Распределение электроэнергии",
 "35.14":"Торговля электроэнергией (энергосбыт)",
 "38.11":"Сбор неопасных отходов (регоператор ТКО)",
 "38.12":"Сбор опасных отходов",
 "38.1":"Сбор отходов (группа)",
}
SECTION_HUMAN = {
 "35":"Обеспечение эл/эн, газом и паром; кондиционирование",
 "36":"Забор, очистка и распределение воды",
 "37":"Сбор и обработка сточных вод",
 "38":"Сбор/обработка/утилизация отходов (ТКО)",
 "68":"Операции с недвижимым имуществом (УК/аренда)",
 "01":"Сельское хозяйство",
 "86":"Здравоохранение (больницы/санатории)",
 "85":"Образование",
 "81":"Обслуживание зданий/территорий (клининг/благоустройство)",
 "41":"Строительство зданий",
 "43":"Специализированные строительные работы",
 "52":"Складирование и вспомогательная транспортная деят-ность",
 "84":"Госуправление / военная безопасность",
 "10":"Пищевая промышленность",
 "87":"Соцобслуживание с проживанием",
 "46":"Оптовая торговля",
 "42":"Строительство инженерных сооружений",
 "47":"Розничная торговля",
 "49":"Транспорт сухопутный/трубопроводный",
 "25":"Металлические изделия",
 "23":"Прочая неметаллическая минеральная продукция",
 "33":"Ремонт/монтаж оборудования",
 "55":"Гостиницы/общепит",
 "16":"Обработка древесины",
 "02":"Лесоводство/лесозаготовки",
 "71":"Архитектура/инжиниринг",
 "72":"Научные исследования",
 "20":"Химическая промышленность",
 "26":"Компьютеры/оптика",
 "82":"Админ/офисная деятельность",
 "30":"Производство прочего транспорта",
 "70":"Консалтинг",
 "27":"Электрооборудование",
 "28":"Машины/оборудование",
 "24":"Металлургия",
 "96":"Прочие персональные услуги",
 "77":"Аренда/лизинг",
 "11":"Напитки",
 "69":"Юридические/бухгалтерские услуги",
 "06":"Добыча нефти/газа",
 "14":"Одежда",
 "29":"Автотранспорт",
 "05":"Добыча угля",
 "07":"Добыча металлических руд",
 "08":"Добыча прочих полезных ископаемых",
 "13":"Текстиль",
 "56":"Общепит",
 "45":"Торговля/ремонт автотранспорта",
 "17":"Бумага",
 "93":"Спорт/отдых",
 "22":"Резина/пластмассы",
 "03":"Рыболовство",
 "64":"Финансовые услуги",
 "95":"Ремонт компьютеров/предметов обихода",
 "90":"Творчество/развлечения",
 "15":"Кожа/обувь",
 "19":"Кокс/нефтепродукты",
 "88":"Соцуслуги без проживания",
 "31":"Мебель",
 "32":"Прочие готовые изделия",
 "39":"Ликвидация загрязнений",
 "94":"Общественные организации",
 "18":"Полиграфия",
 "79":"Туризм",
 "62":"ИТ/ПО",
 "59":"Кино/видео/музыка",
 "91":"Библиотеки/архивы/музеи",
 "50":"Водный транспорт",
 "21":"Фармацевтика",
 "61":"Телеком",
 "60":"Теле/радиовещание",
 "09":"Услуги в добыче полезных ископаемых",
 "80":"Охрана/расследования",
 "75":"Ветеринария",
 "66":"Вспомогательная фин.деятельность",
}

def classify(okved):
    if not okved or not str(okved).strip():
        return ("БЕЗ ОКВЭД", "OTHER", "—", "—")
    o=str(okved).strip()
    if o.startswith("35.30") or o in ("35.30","35.3"):
        return ("Теплоснабжение","TEPLO", OKVED_HUMAN.get(o, "Тепло (35.30)"), SECTION_HUMAN.get("35",""))
    if o.startswith("36.") or o.startswith("37.") or o in ("36.0","37.0"):
        sec=o.split(".")[0]
        return ("Водоснабжение / Водоотведение","VODA", OKVED_HUMAN.get(o, f"Вода/канализация ({o})"), SECTION_HUMAN.get(sec,""))
    if o.startswith("35.22") or o.startswith("35.23") or o.startswith("49.50"):
        return ("Газоснабжение","GAZ", OKVED_HUMAN.get(o, f"Газ ({o})"), SECTION_HUMAN.get(o.split(".")[0],""))
    if o.startswith("35.11") or o.startswith("35.12") or o.startswith("35.13") or o.startswith("35.14"):
        return ("Электроснабжение","ELEC", OKVED_HUMAN.get(o, f"Электричество ({o})"), SECTION_HUMAN.get("35",""))
    if o.startswith("38."):
        return ("ТКО (регоператор)","TKO", OKVED_HUMAN.get(o, f"Отходы ({o})"), SECTION_HUMAN.get("38",""))
    # non-RSO: return section description
    sec=o.split(".")[0]
    return ("НЕ РСО","NON_RSO", OKVED_HUMAN.get(o, ""), SECTION_HUMAN.get(sec, f"Раздел {sec}"))

# Load DaData mapping
inn_map={}
with open(JSONL, encoding="utf-8") as f:
    for line in f:
        j=json.loads(line)
        inn=j["inn"]
        sugg=j.get("response",{}).get("data",{}).get("suggestions",[])
        if sugg:
            d=sugg[0]["data"]
            inn_map[inn] = {
                "okved": d.get("okved"),
                "status": d.get("state",{}).get("status") if d.get("state") else "",
                "status_code": d.get("state",{}).get("code") if d.get("state") else "",
            }
        else:
            inn_map[inn] = {"okved": None, "status": "NO_DATA", "status_code":""}

# Status translation
STATUS_HUMAN={"ACTIVE":"Действующая","BANKRUPT":"Банкрот","LIQUIDATING":"В стадии ликвидации","LIQUIDATED":"Ликвидирована","REORGANIZING":"Реорганизация","NO_DATA":"Нет данных ЕГРЮЛ"}

def annotate():
    wb=openpyxl.load_workbook(SRC_XLSX)
    ws=wb["Лист1"]
    max_col=ws.max_column  # 29 = AC
    header_row=list(ws.iter_rows(min_row=1, max_row=1, values_only=True))[0]
    inn_idx=header_row.index("ИНН")  # 5
    # New columns: after AC (29)
    new_headers=[
        "Категория РСО (по ОКВЭД DaData)",
        "Код категории",
        "ОКВЭД-2 (DaData)",
        "Расшифровка ОКВЭД-2",
        "Раздел ОКВЭД-2",
        "Статус ЕГРЮЛ (DaData)",
        "Признак «мусора»",
        "Пояснение категории",
    ]
    fills={
        "header": PatternFill(start_color="0B5394", end_color="0B5394", fill_type="solid"),
        "TEPLO": PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"), # orange light
        "VODA": PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid"), # blue light
        "GAZ": PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"), # green light
        "ELEC": PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"), # yellow
        "TKO": PatternFill(start_color="E4DFEC", end_color="E4DFEC", fill_type="solid"), # purple
        "NON_RSO": PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"), # gray
        "OTHER": PatternFill(start_color="FCE4EC", end_color="FCE4EC", fill_type="solid"),
    }
    font_header=Font(color="FFFFFF", bold=True, size=8, name="Calibri")
    font_cell=Font(size=8, name="Calibri")
    font_bold=Font(size=8, name="Calibri", bold=True)
    thin=Border(left=Side(style='thin', color='B4C6E7'), right=Side(style='thin', color='B4C6E7'), top=Side(style='thin', color='B4C6E7'), bottom=Side(style='thin', color='B4C6E7'))
    align_c=Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_l=Alignment(horizontal='left', vertical='center', wrap_text=True)

    start_col=max_col+1  # 30 = AD
    # widths
    widths=[32,10,12,40,28,18,18,44]
    for i, h in enumerate(new_headers):
        col=start_col+i
        c=ws.cell(row=1, column=col, value=h)
        c.fill=fills["header"]
        c.font=font_header
        c.alignment=align_c
        c.border=thin
        ws.column_dimensions[get_column_letter(col)].width=widths[i]
    ws.row_dimensions[1].height=42
    ws.freeze_panes="A2"
    # adjust autofilter
    ws.auto_filter.ref=ws.dimensions

    # Data rows
    rows=list(ws.iter_rows(values_only=True))
    for r_idx in range(2, ws.max_row+1):
        # INN is col 6 (F)
        inn_cell=ws.cell(row=r_idx, column=6)  # F
        inn=str(inn_cell.value).strip() if inn_cell.value else ""
        info=inn_map.get(inn, {"okved": None, "status":"", "status_code":""})
        okved=info.get("okved")
        status=info.get("status") or ""
        status_h=STATUS_HUMAN.get(status, status)
        cat, code, desc, section = classify(okved)
        # Признак мусора
        if code=="NON_RSO":
            flag="Мусор: не РСО"
        elif status in ("LIQUIDATED","LIQUIDATING","BANKRUPT"):
            flag={"LIQUIDATED":"Ликвидирована","LIQUIDATING":"Ликвидируется","BANKRUPT":"Банкрот"}[status]
            # keep category but flag
            if code=="NON_RSO":
                flag="Мусор + "+flag
        else:
            flag="ОК"
        # Пояснение
        if code=="TEPLO":
            note="Ресурс: тепло/горячая вода/котельная/теплосеть (ОКВЭД 35.30*)"
        elif code=="VODA":
            note="Ресурс: питьевая/техническая вода + водоотведение/канализация (36.*/37.*)"
        elif code=="GAZ":
            note="Ресурс: природный/сжиженный газ по сетям/трубопроводам (35.22/35.23/49.50*)"
        elif code=="ELEC":
            note="Ресурс: электроэнергия — генерация/сети/сбыт (35.11*/35.12*/35.13/35.14)"
        elif code=="TKO":
            note="Ресурс: ТКО — региональный оператор по обращению с отходами (38.11/38.12)"
        else:
            # non-RSO: explain
            if okved:
                sec=okved.split(".")[0]
                note=f"НЕ является РСО: заявлена роль РСО в ГИС ЖКХ, но основной ОКВЭД {okved} — {section or 'раздел '+sec}."
            else:
                note="НЕ является РСО: нет ОКВЭД в ЕГРЮЛ."

        values=[cat, code, okved or "", desc, section, f"{status_h} ({status})" if status else "", flag, note]
        fill=fills.get(code, fills["OTHER"])
        for i, v in enumerate(values):
            col=start_col+i
            c=ws.cell(row=r_idx, column=col, value=v)
            c.font=font_bold if i in (0,1,6) else font_cell
            c.alignment=align_c if i in (0,1,2,5,6) else align_l
            c.border=thin
            c.fill=fill
            # highlight bad status even inside correct category
            if i==5 and status in ("LIQUIDATING","LIQUIDATED","BANKRUPT"):
                c.font=Font(color="FF0000", bold=True, size=8, name="Calibri")
            if i==6 and flag.startswith("Мусор"):
                c.font=Font(color="FF0000", bold=True, size=8, name="Calibri")
        ws.row_dimensions[r_idx].height=14

    # Add summary row at bottom? Let's add a second sheet "Сводка"
    # Create сводка sheet if not exists
    if "Сводка" in wb.sheetnames:
        wb.remove(wb["Сводка"])
    summary=wb.create_sheet("Сводка", 0)
    summary.sheet_properties.tabColor="0B5394"
    title="РСО — механизм разбивки по назначению (на основе ОКВЭД-2 из DaData / ЕГРЮЛ)"
    summary.merge_cells("A1:H1")
    c=summary.cell(row=1, column=1, value=title)
    c.font=Font(size=12, bold=True, color="0B5394", name="Calibri")
    c.alignment=Alignment(horizontal='center', vertical='center')
    summary.row_dimensions[1].height=22

    summary["A2"]="Источник: РСО_DaData.xlsx (7008 записей, поле ГИС ЖКХ «Функция» = у всех 'Ресурсоснабжающая организация' — неинформативно) + РСО.jsonl (DaData suggestions[0].data.okved, state.status, проверка 2026-11-16). Метод: классификация по первой части ОКВЭД-2."
    summary.merge_cells("A2:H2")
    summary["A2"].font=Font(size=7, italic=True, color="595959", name="Calibri")
    summary["A2"].alignment=Alignment(wrap_text=True, vertical='center')
    summary.row_dimensions[2].height=30

    headers2=["Код","Категория","Кол-во","Доля","Главные ОКВЭДы в группе","Что это на практике","Что делать","Цвет"]
    col_widths2=[8,26,8,8,42,40,34,6]
    for i,h in enumerate(headers2):
        col=i+1
        c=summary.cell(row=3, column=col, value=h)
        c.fill=fills["header"]
        c.font=font_header
        c.alignment=align_c
        c.border=thin
        summary.column_dimensions[get_column_letter(col)].width=col_widths2[i]
    summary.row_dimensions[3].height=28

    # Calculate stats
    from collections import Counter
    cnt=Counter()
    for r in range(2, ws.max_row+1):
        code=ws.cell(row=r, column=start_col+1).value  # code
        cnt[code]+=1
    total=ws.max_row-1
    data=[
        ("TEPLO","Теплоснабжение (котельные/ТС)", "35.30* — 35.30.14 котельные (1458), 35.30 группа (617), 35.30.3 распределение (107), 35.30.2 передача (102), 35.30.1 производство (142), 35.30.4 сети (60)", "Котельные, ТЭЦ, теплосети, передача/распределение тепла и гор.воды", "Оставить, разделить по подкатегориям: котельные / сети / производство", fills["TEPLO"]),
        ("VODA","Водоснабжение / Водоотведение", "36.00.2 распределение воды (739), 36.00.1 забор/очистка (448), 36.00 группа (360), 37.00 сточные воды (308)", "Водоканалы, водопроводы, очистные, канализация", "Оставить; 37.00 → отдельно 'Водоотведение', остальное 'Водоснабжение'", fills["VODA"]),
        ("GAZ","Газоснабжение", "35.22 распределение газа (59), 35.23 торговля газом (36), 49.50.21 транспортирование газа (11)", "Облгазы, горгазы, Газпром газораспределение", "Оставить как есть", fills["GAZ"]),
        ("ELEC","Электроснабжение", "35.14 сбыт (78), 35.12.1 сети (36), 35.11.1 ТЭС (51), 35.11 (29)", "Электросети, сбытовые компании, ТЭЦ/ГРЭС", "Оставить; сбыт (35.14) можно отделить от сетей", fills["ELEC"]),
        ("TKO","ТКО (регоператор)", "38.11 сбор неопасных отходов (37), 38.12 и др.", "Региональные операторы по обращению с отходами", "Проверить: в ГИС ЖКХ роль РСО ставят регоператорам ТКО — вынести в отдельный тип", fills["TKO"]),
        ("NON_RSO","НЕ РСО — мусорные записи", "68.32 УК (277), 86.90.4 санатории (84), 41.20 стройка (62), 01.41 молоко/КРС (94), 81.29 клининг (51), 68.20 аренда (83) и т.д. (см. лист)", "УК, школы, санатории, заводы, совхозы, стройки, фермы — заявлены как РСО ошибочно", "Исключить из РСО-рассылок/реестра; вернуть на верификацию в ГИС ЖКХ или перенести в УК/прочие", fills["NON_RSO"]),
    ]
    r0=4
    for idx,(code,cat,okveds_practice,practice,action,fill) in enumerate(data):
        n=cnt.get(code,0)
        share=f"{n/total:.1%}" if total else ""
        row=[code,cat,n,share,okveds_practice,practice,action,""]
        for col, val in enumerate(row,1):
            c=summary.cell(row=r0+idx, column=col, value=val)
            c.font=font_cell
            c.alignment=align_l if col in (2,5,6,7) else align_c
            c.border=thin
            c.fill=fill
            if col==4:
                c.number_format='0.0%'
            if col==1:
                c.font=font_bold
        summary.row_dimensions[r0+idx].height=44
        summary.cell(row=r0+idx, column=8).fill=fill

    # Итог row
    rr=r0+len(data)
    summary.cell(row=rr, column=1, value="ИТОГО").font=font_bold
    summary.cell(row=rr, column=3, value=total).font=font_bold
    summary.cell(row=rr, column=4, value="100%").font=font_bold
    for col in range(1,9):
        summary.cell(row=rr, column=col).border=thin
        summary.cell(row=rr, column=col).alignment=align_c
    summary.row_dimensions[rr].height=18

    # Notes block
    nr=rr+2
    notes=[
        "Вывод 1 — поле ГИС ЖКХ «Функция» бесполезно: у всех 7008 строк одно значение 'Ресурсоснабжающая организация'. Разбивки Газ/Вода/Тепло/Свет там нет.",
        "Вывод 2 — единственный рабочий механизм — основной ОКВЭД-2 из ЕГРЮЛ (DaData okved). Он определяет, чем организация занимается на самом деле.",
        "Вывод 3 — правило: 35.30* → тепло, 36.*/37.* → вода/канализация, 35.22/35.23/49.50* → газ, 35.11*/35.12*/35.13/35.14 → электричество, 38.* → ТКО. Всё остальное → НЕ РСО.",
        "Вывод 4 — ~30% (2110 из 7008) — мусор: не РСО вообще. Топ-мусор: 68.32* УК (277), 86*/87* санатории/больницы (196), 01.* сельское хозяйство (214), 41-43 стройка (213), 85* образование (106), 81* клининг (90), 68.20 аренда (83) и др. Примеры: '103 АРСЕНАЛ' (25.40 оружие), 'БРОНЕТАНКОВЫЙ ЗАВОД' (30.40), 'УЧЕБНЫЙ ЦЕНТР ЛЕСОТЕХШКОЛА' (85.42), 'ДОМ ПИСАТЕЛЕЙ ПЕРЕДЕЛКИНО' (90.04), 'САНАТОРИЙ КРАСНЫЙ БОР' (86.90).",
        "Вывод 5 — даже 'говорящие' названия врут: среди ~490 'ВОДОКАНАЛ' только ~60% имеют водный ОКВЭД; остальные — тепловые (35.30.14 ТЕПЛОВОДОКАНАЛ), арендные (68.20), строительные (42.21) и даже ТКО (38.11). Доверять только ОКВЭДу.",
        "Вывод 6 — статусы ЕГРЮЛ: ACTIVE 6211 (88.6%), BANKRUPT 522 (7.4%), LIQUIDATING 130, LIQUIDATED 112, REORGANIZING 33. Их тоже помечаем в колонке 'Признак' — банкроты/ликвидируемые исключать из рассылок.",
        "Рекомендация: в РСО_DaData_категоризировано.xlsx фильтруйте по 'Код категории' = TEPLO/VODA/GAZ/ELEC/TKO для целевых рассылок; NON_RSO — в отдельный список на проверку; BANKRUPT/LIQUIDATED — в стоп-лист.",
    ]
    for i, txt in enumerate(notes):
        summary.merge_cells(f"A{nr+i}:H{nr+i}")
        c=summary.cell(row=nr+i, column=1, value=txt)
        c.font=Font(size=7, name="Calibri", color="404040")
        c.alignment=Alignment(wrap_text=True, vertical='center')
        summary.row_dimensions[nr+i].height=22

    # Print settings
    ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.page_setup.orientation='landscape'
    ws.page_setup.fitToWidth=1
    summary.sheet_properties.pageSetUpPr.fitToPage=True
    summary.page_setup.orientation='landscape'
    summary.page_setup.fitToWidth=1
    ws.print_options.horizontalCentered=True
    summary.print_options.horizontalCentered=True

    wb.save(DST_XLSX)
    print(f"Saved {DST_XLSX} rows={ws.max_row} cols={ws.max_column}")
    # also print counts
    for k,v in cnt.items():
        print(k, v)

if __name__=="__main__":
    annotate()
