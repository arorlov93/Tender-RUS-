"""Камчатка 32616346943: спецификация для запроса КП по оставшимся позициям (краны, затворы, приводы) — дословно из Прил. 1 к ТТ (data/kamch/tz_verbatim.json).
С аргументом — добавляет лист «Заводы» из проверенных результатов поиска (rem_krany.json, rem_zatvory.json, только verified=true)."""
import json, os, sys, openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

T = {(t["table"], t["n_table"]): t for t in json.load(open("data/kamch/tz_verbatim.json"))}
R = json.load(open("data/kamch/remaining_2310.json"))
thin = Side(style="thin", color="BFBFBF"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin); FILL = PatternFill("solid", fgColor="FFF2CC")
COLS = ["Наименование", "ГОСТ/ТУ", "Ду, мм", "Проход", "Давление, МПа", "Температура, °С", "Среда", "Материал корпуса", "Климатическое исполнение", "Присоединение",
        "Приводное устройство: наименование", "Приводное устройство: технические характеристики", "Установочное положение", "Материал штока, шара, уплотнений", "Габариты", "Примечание"]
FACT = ["Предлагаемый товар (марка, артикул)", "Страна производства", "Номер записи в реестре Минпромторга", "Цена за шт с НДС, ₽", "Срок изготовления/поставки", "Соответствует ТЗ полностью? (да/нет, отличия)"]


def head(ws, row):
    for c in ws[row]:
        if c.value is not None: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")


wb = openpyxl.Workbook(); first = wb.active
SHEETS = [("Краны", "Краны"), ("Затворы и обратные", "Затворы"), ("Приводы и прочее", "Приводы и прочее")]
for title, g in SHEETS:
    ws = wb.create_sheet(title)
    ws.append([f"Запрос КП: {title.lower()} — АО «Камчатскэнерго» (поставка г. Петропавловск-Камчатский). Требования — дословно из Приложения 1 к техническим требованиям заказчика."])
    ws["A1"].font = Font(bold=True, size=12)
    ws.append(["Нацрежим ПП 1875 (ограничение): товар российского производства с записью в реестре Минпромторга. Изготовление — в год поставки или предыдущий. Жёлтые поля заполняет завод."])
    ws.append(["№", "Таблица / № в ТЗ", "Кол-во", "Ед."] + COLS + FACT); head(ws, 3)
    k = 0
    for r in [r for r in R if r["group"] == g]:
        t = T[(r["table"], r["n"])]; k += 1
        ws.append([k, f'{r["table"]} / {r["n"]}', t["qty"], t["unit"]] + [t[c] for c in COLS] + [None] * len(FACT))
        X = ws.max_row
        for j in range(5 + len(COLS), 5 + len(COLS) + len(FACT)): ws.cell(X, j).fill = FILL
        for c in ws[X]: c.border = BOX; c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[X].height = 75
    W = [4, 10, 7, 5, 32, 22, 6, 12, 8, 14, 18, 18, 10, 18, 16, 22, 14, 26, 16, 18, 26, 14, 18, 14, 14, 22]
    for i, w in enumerate(W, 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.row_dimensions[3].height = 60; ws.freeze_panes = "F4"
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0; ws.sheet_properties.pageSetUpPr.fitToPage = True

# лист «Заводы» — только verified=true
src = [f for f in ("data/kamch/rem_krany.json", "data/kamch/rem_zatvory.json") if os.path.exists(f)]
first.title = "Заводы"
first.append(["Заводы-производители для оставшихся позиций (краны, затворы, приводы) — только проверенные: производство, номенклатура и контакты подтверждены на официальном сайте"])
first["A1"].font = Font(bold=True, size=12)
first.append(["Проверено: 1) своё производство; 2) делает именно эти позиции (ссылка на каталог); 3) телефон и e-mail совпадают минимум в двух местах официального сайта."])
first.append(["№", "Завод", "Город", "Какие позиции закрывает (таблица / №)", "Что делает", "Реестр Минпромторга", "Телефон", "E-mail", "Сайт", "Где проверен контакт", "Каталог (ссылка)", "Отправлен запрос", "Ответ"])
head(first, 3)
n = 0
for f in src:
    for x in json.load(open(f)):
        if not x.get("verified"): continue
        n += 1
        first.append([n, x["company"], x.get("plant_city"), ", ".join(x.get("covers_tt") or [])[:300], str(x.get("covers_note") or "")[:200], str(x.get("gisp") or "")[:120],
                      x.get("phone"), x.get("email"), x.get("site"), str(x.get("contact_check") or "")[:160], (x.get("catalog_urls") or [""])[0], None, None])
        for c in first[first.max_row]: c.border = BOX; c.alignment = Alignment(wrap_text=True, vertical="top")
        first.row_dimensions[first.max_row].height = 60
if not n: first.append(["", "поиск заводов идёт — лист заполнится после проверки"])
for i, w in enumerate([4, 30, 14, 30, 36, 22, 24, 24, 16, 30, 30, 12, 16], 1): first.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
first.row_dimensions[3].height = 40; first.freeze_panes = "C4"
first.page_setup.orientation = "landscape"; first.page_setup.fitToWidth = 1; first.page_setup.fitToHeight = 0; first.sheet_properties.pageSetUpPr.fitToPage = True
# лист «Покрытие»: по каждой позиции — какие проверенные заводы её закрывают
V = [x for f in src for x in json.load(open(f)) if x.get("verified")]
cv = wb.create_sheet("Покрытие", 1)
cv.append(["Таблица / № в ТЗ", "Позиция ТЗ", "Кол-во", "Проверенных заводов", "Какие заводы"]); head(cv, 1)
NOTE = {"1.3/135": "не закрыта: у проверенных заводов нет нерж. муфтового Ду15 на Ру63 — искать дополнительно",
        "1.3/32": "только частичный аналог (Гранлок RD16, не аксиальный) — согласовать эквивалент с заказчиком", "1.3/33": "только частичный аналог (Гранлок RD16) — согласовать",
        "1.3/34": "только частичный аналог (Гранлок RD16) — согласовать", "1.6/93": "аналог ESBE — МЭОФ ЗЭиМ крупнее, проверить посадку на клапан",
        "1.3/150": "точного аналога 32нж910р нет — исполнение (морская вода, КОФ, электропривод) согласовать с заводом"}
for r in R:
    tt = f'{r["table"]}/{r["n"]}'; t = T[(r["table"], r["n"])]
    fs = [x["company"].split("(")[0].strip() for x in V if tt in (x.get("covers_tt") or [])]
    cv.append([tt, t["Наименование"], t["qty"], len(fs), "; ".join(fs) + (f" — {NOTE[tt]}" if tt in NOTE else "")])
    cv.cell(cv.max_row, 4).fill = PatternFill("solid", fgColor="F8CBAD" if not fs else "FFEB9C" if len(fs) == 1 or tt in NOTE else "C6EFCE")
    for c in cv[cv.max_row]: c.border = BOX; c.alignment = Alignment(wrap_text=True, vertical="top")
for i, w in enumerate([12, 50, 8, 12, 90], 1): cv.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
cv.freeze_panes = "C2"; cv.auto_filter.ref = f"A1:E{cv.max_row}"
print("без заводов:", [c[0].value for c in cv.iter_rows(min_row=2) if c[3].value == 0])
wb.save("../KAMCHATKA_REMAINING_ZAVODY_SPEC.xlsx")
print("заводов", n, "| позиций", {g: sum(1 for r in R if r["group"] == g) for _, g in SHEETS})
