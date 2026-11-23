#!/usr/bin/env python3
"""
Добавляет 4 колонки со склонением в Администрациях_городов_обогащенная_DaData.xlsx:
1. Должность (дательный, ВЕРХНИЙ РЕГИСТР) — ГЛАВЕ ГОРОДА АБАКАНА
2. ФИО полностью (дательный) — Лемину Алексею Викторовичу
3. Фамилия И.О. (дательный) — Лемину А.В.
4. Обращение — Уважаемый Алексей Викторович!
"""
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from petrovich.main import Petrovich
from petrovich.enums import Case, Gender
import pymorphy3

PATH = "/home/user/gov/Администрации_городов_обогащенная_DaData.xlsx"
petro = Petrovich()
morph = pymorphy3.MorphAnalyzer()

HEADER_FILL_DECL = PatternFill(start_color="0B5394", end_color="0B5394", fill_type="solid")  # синий для склонений
HEADER_FONT = Font(color="FFFFFF", bold=True, size=9, name="Calibri")
THIN_BORDER = Border(
    left=Side(style='thin', color='B4C6E7'),
    right=Side(style='thin', color='B4C6E7'),
    top=Side(style='thin', color='B4C6E7'),
    bottom=Side(style='thin', color='B4C6E7')
)

def detect_gender(fio: str):
    if not fio or not fio.strip():
        return Gender.MALE
    fio = fio.strip()
    parts = fio.split()
    if len(parts) >= 3:
        patronymic = parts[2].lower()
        if patronymic.endswith("вна") or patronymic.endswith("чна") or patronymic.endswith("шна") or patronymic.endswith("овна") or patronymic.endswith("евна") or patronymic.endswith("ична") or patronymic.endswith("ьна") or patronymic.endswith("инична"):
            return Gender.FEMALE
        if patronymic.endswith("ич") or patronymic.endswith("вич") or patronymic.endswith("льич") or patronymic.endswith("мич"):
            return Gender.MALE
    if len(parts) >= 2:
        firstname = parts[1].lower()
        # try pymorphy
        try:
            p = morph.parse(firstname)[0]
            if 'femn' in p.tag:
                return Gender.FEMALE
            if 'masc' in p.tag:
                return Gender.MALE
        except:
            pass
        if firstname.endswith("а") or firstname.endswith("я"):
            if firstname in ("никита", "илья", "кузьма", "фома", "лука", "миха"):
                return Gender.MALE
            return Gender.FEMALE
    # fallback check surname ending for female: -ва, -на, -ая
    if len(parts) >= 1:
        surname = parts[0].lower()
        if surname.endswith("ва") or surname.endswith("на") or surname.endswith("ая") or surname.endswith("ова") or surname.endswith("ева"):
            # could be female, but male also -ов, need more precise: female surname often ends with а, male with consonant
            # if surname ends with а -> likely female
            if surname.endswith("а"):
                # check patronymic already handled, so approximate
                pass
    return Gender.MALE

def decline_fio_full(fio: str):
    if not fio or not fio.strip():
        return ""
    fio = fio.strip()
    parts = fio.split()
    if len(parts) < 2:
        return fio
    gender = detect_gender(fio)
    surname = parts[0]
    firstname = parts[1] if len(parts) > 1 else ""
    patronymic = parts[2] if len(parts) > 2 else ""
    try:
        d_surname = petro.lastname(surname, Case.DATIVE, gender) if surname else ""
    except:
        d_surname = surname
    try:
        d_first = petro.firstname(firstname, Case.DATIVE, gender) if firstname else ""
    except:
        d_first = firstname
    try:
        d_mid = petro.middlename(patronymic, Case.DATIVE, gender) if patronymic else ""
    except:
        d_mid = patronymic
    res = " ".join([x for x in [d_surname, d_first, d_mid] if x])
    return res

def fio_to_initials_dative(fio: str):
    """Лемину А.В. — дательный фамилия + инициалы"""
    if not fio or not fio.strip():
        return ""
    parts = fio.strip().split()
    if len(parts) < 2:
        return fio
    gender = detect_gender(fio)
    surname = parts[0]
    firstname = parts[1] if len(parts) > 1 else ""
    patronymic = parts[2] if len(parts) > 2 else ""
    try:
        d_surname = petro.lastname(surname, Case.DATIVE, gender) if surname else ""
    except:
        d_surname = surname
    # initials from original names (upper + dot)
    ini_f = (firstname[0].upper() + ".") if firstname else ""
    ini_m = (patronymic[0].upper() + ".") if patronymic else ""
    initials = ini_f + ini_m
    # format "Лемину А.В." — surname + space + initials
    if initials:
        return f"{d_surname} {initials}"
    else:
        return d_surname

def respectful_address(fio: str):
    """Уважаемый Алексей Викторович!"""
    if not fio or not fio.strip():
        return ""
    parts = fio.strip().split()
    if len(parts) < 2:
        return f"Уважаемый {fio}!"
    gender = detect_gender(fio)
    firstname = parts[1]
    patronymic = parts[2] if len(parts) > 2 else ""
    respect = "Уважаемая" if gender == Gender.FEMALE else "Уважаемый"
    if patronymic:
        return f"{respect} {firstname} {patronymic}!"
    else:
        return f"{respect} {firstname}!"

def decline_position(phrase: str):
    if not phrase or not phrase.strip():
        return ""
    phrase = phrase.strip()
    words = phrase.split()
    out = []
    for w in words:
        # handle hyphenated? keep as is for simplicity
        # Strip punctuation
        prefix = ""
        suffix = ""
        core = w
        # keep quotes etc — not needed
        low = core.lower()
        p = morph.parse(low)[0]
        if 'nomn' in p.tag:
            inf = p.inflect({"datv"})
            if inf:
                # preserve case: original is upper
                if w.isupper():
                    out.append(inf.word.upper())
                elif w[0].isupper():
                    out.append(inf.word.capitalize())
                else:
                    out.append(inf.word)
            else:
                out.append(w)
        else:
            out.append(w)
    # preserve all-upper if input was all upper
    res = " ".join(out)
    if phrase.isupper():
        res = res.upper()
    return res

def process_workbook():
    wb = openpyxl.load_workbook(PATH)
    # headers for new 4 cols
    new_headers = [
        "1. Должность (дат. п., ВЕРХНИЙ РЕГИСТР) — ГЛАВЕ ГОРОДА АБАКАНА (пример)",
        "2. Кому: ФИО полностью (дат. п.) — Лемину Алексею Викторовичу",
        "3. Кому: Фамилия И.О. (дат. п.) — Лемину А.В.",
        "4. Обращение — Уважаемый Алексей Викторович!"
    ]
    sheets_to_process = [s for s in wb.sheetnames if s != "Оглавление"]
    for sheet_name in sheets_to_process:
        ws = wb[sheet_name]
        print(f"Обработка листа {sheet_name} ... max_col={ws.max_column} max_row={ws.max_row}")
        rows = list(ws.iter_rows(values_only=True))
        header = list(rows[0])
        # find indexes
        try:
            fio_idx = header.index("Руководитель ФИО (DaData)")
            dol_idx = header.index("Должность руководителя (DaData)")
        except ValueError as e:
            print(f"  пропущен лист {sheet_name}: нет нужных колонок {e}")
            continue
        # add new headers at end
        start_col = ws.max_column + 1
        for i, h in enumerate(new_headers):
            col = start_col + i
            cell = ws.cell(row=1, column=col, value=h)
            cell.fill = HEADER_FILL_DECL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = THIN_BORDER
        # set width
        widths = [38, 36, 22, 34]
        for i, w in enumerate(widths):
            col_letter = get_column_letter(start_col + i)
            ws.column_dimensions[col_letter].width = w
        ws.row_dimensions[1].height = 38

        # iterate data rows
        for r_idx in range(2, ws.max_row + 1):
            fio = ws.cell(row=r_idx, column=fio_idx + 1).value
            dol = ws.cell(row=r_idx, column=dol_idx + 1).value
            # normalize empty
            fio_str = str(fio).strip() if fio else ""
            dol_str = str(dol).strip() if dol else ""

            # compute
            if dol_str:
                decl_dol = decline_position(dol_str)
            else:
                decl_dol = ""

            if fio_str:
                fio_dative = decline_fio_full(fio_str)
                fio_io = fio_to_initials_dative(fio_str)
                respect = respectful_address(fio_str)
            else:
                fio_dative = ""
                fio_io = ""
                respect = ""

            # write
            vals = [decl_dol, fio_dative, fio_io, respect]
            for i, v in enumerate(vals):
                col = start_col + i
                cell = ws.cell(row=r_idx, column=col, value=v)
                cell.alignment = Alignment(vertical='center', wrap_text=True)
                cell.border = THIN_BORDER
                cell.font = Font(size=9, name="Calibri")
                # alternate row color? keep original style for old cols, for new cols use light blue alt
                if r_idx % 2 == 0:
                    # light gray-blue for new cols
                    cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
            # height
            ws.row_dimensions[r_idx].height = 16

        # freeze and filter update
        ws.freeze_panes = "A2"
        # extend autofilter
        try:
            ws.auto_filter.ref = ws.dimensions
        except:
            pass
        print(f"  добавлено 4 колонки, всего {ws.max_column}")

    wb.save(PATH)
    print(f"Сохранено {PATH}")
    # print stats
    wb2 = openpyxl.load_workbook(PATH, read_only=True)
    for sn in sheets_to_process:
        ws = wb2[sn]
        print(sn, ws.dimensions, ws.max_column)

if __name__ == "__main__":
    process_workbook()
