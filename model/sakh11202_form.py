"""Сахалинэнерго лот 2 № 11202 (НМЦ 16 153 216,67 ₽ без НДС, подача 05.11.2026): заполнение ОФИЦИАЛЬНОЙ формы 3 «КП и структура НМЦ»
(data/sakh/docs11202/…/Форма 3 КП и Структура НМЦ.xlsx) — столбцы 6–8 (страна, реестр ПП 719 и № записи, производитель) и 10 (цена ед. без НДС).
Режим по всем 87 позициям — «ограничение» (ПП 1875). Номера — из выгрузки ГИСП 01.11.2026 (действующие записи), держатель = производитель.
Цена ед. = НМЦ ед. × (1 − DISC), DISC = 10 % по умолчанию (argv[1] — другая скидка). Правка XML — проверки данных формы сохраняются.
→ ../SAKHALIN_11202_ZAYAVKA/02a_КП_форма3_заполнено.xlsx + 02b_Реестр_сверка.xlsx"""
import json, re, sys, os, zipfile, openpyxl
from lxml import etree
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from tidy import tidy

DISC = float(sys.argv[1]) if len(sys.argv) > 1 else 0.10
SRC = "data/sakh/docs11202/Приложение № 4 - Формы документов заявки/Форма 3 КП и Структура НМЦ.xlsx"
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
GF = "data/sakh/gisp_records_11202.json"
NEED = {"10655491", "10826766", "10895023", "10636782", "10457924", "10232727", "10232731", "10698937", "10618662", "10618660", "10618658", "10895294",
        "10357202", "10837159", "10382110", "10811565", "10507445", "10507446", "10647852", "10763482", "10680853"}
if len(sys.argv) > 2:  # извлечь записи из распарсенной выгрузки ГИСП (jsonl)
    G = {}
    for f in sys.argv[2:]:
        for l in open(f):
            r = json.loads(l)
            if r["G"] in NEED: G.setdefault(r["G"], []).append(r)
    json.dump(G, open(GF, "w"), ensure_ascii=False, indent=1)
G = json.load(open(GF))
act = lambda n: [r for r in G.get(n, []) if not r.get("K") and (r.get("J") or "9999") >= "2026-11-05"]
def short(a):
    a = "".join(("«" if i == 0 or a[i - 1] in " «" else "»") if ch == '"' else ch for i, ch in enumerate(a)); a += "»" * (a.count("«") - a.count("»"))
    for k, v in (("ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ", "ООО"), ("ОТКРЫТОЕ АКЦИОНЕРНОЕ ОБЩЕСТВО", "ОАО"), ("АКЦИОНЕРНОЕ ОБЩЕСТВО", "АО")): a = a.replace(k, v)
    return re.sub(r"\s+", " ", a).strip()

# ---- позиции формы
wb0 = openpyxl.load_workbook(SRC); ws0 = wb0.active
POS = []; part = 0
for r in range(15, ws0.max_row + 1):
    q, name = ws0.cell(r, 17).value, ws0.cell(r, 18).value
    if name and str(name).strip().upper().startswith("ПАРТИЯ"): part += 1; continue
    if name and q and str(q).strip(".").isdigit():
        POS.append(dict(row=r, tt=f"{part}/{int(str(q).strip('.'))}", name=str(name), qty=float(ws0.cell(r, 22).value), nmc=float(ws0.cell(r, 25).value)))
dn = lambda s: (lambda m: int(m.group(1)) if m else None)(re.search(r"(?:Ду|DN)\s*(\d+)", s))

# ---- правило: (регэксп по наименованию, № записи, комментарий / эквивалент)
RULES = [
    (r"LD КШЦП|КШ\.Ц\.Ф|LD Ду 80", "10655491", ""),
    (r"11б27п1|VALTEC BASE|Стандарт 1/2.{0,5}бабочка|ГАЛЛОП Практик|Кран шаровый с характеристиками не ниже", "10826766", "эквивалент: кран шаровой латунный LD Pride (ЛД Прайд) по характеристикам ТЗ"),
    (r"15б1п", "10895023", ""),
    (r"30ч6бр Ду 400", "10457924", "задвижка чугунная ЗЛЧ DN400 PN10 (Арма-Пром)"),
    (r"30ч6бр", "10636782", ""),
    (r"30с64нж|30с41нж|30лс964нж", "10232727", "задвижка клиновая ЗКС (Арма-Пром, ТМ САЗ)"),
    (r"15с65нж", "10232731", "клапан запорный КЗ (Арма-Пром)"),
    (r"15кч18п", "10698937", "клапан запорный ПТЗ (эквивалент 15кч18п)"),
    (r"19ч16бр Ду 250", "10618662", "эквивалент: обратный клапан Гранлок ЗОП DN250 PN16"),
    (r"обратный чугун шаровой.*Ду 150", "10618660", "эквивалент: обратный клапан Гранлок ЗОП DN150"),
    (r"обратный чугун шаровой.*Ду 100", "10618658", "эквивалент: обратный клапан Гранлок ЗОП DN100"),
    (r"16ч42р", "10895294", "клапан обратный подъёмный ПЗТА.КО (16ч42р)"),
    (r"ЗПВС.*-100-", "10357202", ""),
    (r"ЗПВС", "10837159", ""),
    (r"МЭОФ", "10382110", ""),
    (r"ВДХ|15ч75п2|футерованный|15п67п|15ч93эм", "10811565", "эквивалент: клапан запорный 15нж65нж из коррозионностойкой стали (ВАК) — футерованных/мембранных с записью ПП 719 нет"),
    (r"Кран шаровой запорный разборный", None, "ЯРДОС"),
    (r"мнометрический КШ250", "10507445", "эквивалент: кран шаровой ЗАРД высокого давления (PN до 320) вместо КШ250 Гималаи"),
    (r"Клапан электромагнитный", "10763482", "эквивалент: клапан электромагнитный КЭО (DN 3–150, PN до 800 кгс/см²) — подтвердить датчик положения затвора и присоединение"),
    (r"Сепаратор грязи", "10680853", "эквивалент: грязевик LD вместо сепаратора Flamco Clean Smart"),
    (r"Классик ФБ39", "10647852", "эквивалент: кран шаровой стальной ALSO DN50 (ФОБОС «Классик» записи ПП 719 нет)"),
]
OUT = {}
for p in POS:
    for rx, num, note in RULES:
        if re.search(rx, p["name"], re.I):
            if note == "ЯРДОС": num = "10507445" if (dn(p["name"]) or 100) <= 100 else "10507446"; note = "кран шаровой ЗАРД (ИК Энерпред-Ярдос)"
            a = act(num)
            if a: OUT[p["tt"]] = dict(num=num, man=f'{short(a[0]["A"])} (ИНН {a[0]["B"]})', prod=a[0]["L"], valid=a[0].get("J"), note=note)
            break

# ---- запись в XML формы
z = zipfile.ZipFile(SRC); root = etree.fromstring(z.read("xl/worksheets/sheet1.xml"))
rows = {int(r.get("r")): r for r in root.find(f"{{{NS}}}sheetData").findall(f"{{{NS}}}row")}
cn = lambda ref: sum((ord(ch) - 64) * 26 ** i for i, ch in enumerate(reversed("".join(c for c in ref if c.isalpha()))))
def cell(row, ref):
    c = row.find(f'{{{NS}}}c[@r="{ref}"]')
    if c is None:
        c = etree.SubElement(row, f"{{{NS}}}c"); c.set("r", ref)
        cs = sorted(row.findall(f"{{{NS}}}c"), key=lambda e: cn(e.get("r")))
        for e in cs: row.remove(e)
        for e in cs: row.append(e)
    for ch in list(c): c.remove(ch)
    return c
def put_s(row, ref, text):
    c = cell(row, ref); c.set("t", "inlineStr"); etree.SubElement(etree.SubElement(c, f"{{{NS}}}is"), f"{{{NS}}}t").text = text
def put_n(row, ref, v):
    c = cell(row, ref)
    if "t" in c.attrib: del c.attrib["t"]
    etree.SubElement(c, f"{{{NS}}}v").text = repr(round(v, 2))
total = 0
for p in POS:
    row = rows[p["row"]]; price = round(p["nmc"] * (1 - DISC), 2); p["price"] = price; total += price * p["qty"]
    put_n(row, f'L{p["row"]}', price)
    o = OUT.get(p["tt"])
    if o:
        put_s(row, f'H{p["row"]}', "Россия"); put_s(row, f'I{p["row"]}', f'Реестр российской промышленной продукции, реестровая запись № {o["num"]}'); put_s(row, f'J{p["row"]}', o["man"])
OD = "../SAKHALIN_11202_ZAYAVKA"; os.makedirs(OD, exist_ok=True)
with zipfile.ZipFile(f"{OD}/02a_КП_форма3_заполнено.xlsx", "w", zipfile.ZIP_DEFLATED) as w:
    for it in z.infolist():
        w.writestr(it, etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True) if it.filename == "xl/worksheets/sheet1.xml" else z.read(it.filename))

# ---- таблица сверки
thin = Side(style="thin", color="BFBFBF"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Сверка реестра"
ws.append([f"Сахалинэнерго лот 2 № 11202: форма 3 заполнена — цена = НМЦ ед. × {1 - DISC:.2f}, итого {total:,.2f} ₽ без НДС (НМЦ 16 153 216,67). Номера ПП 719 — действующие записи ГИСП (выгрузка 01.11.2026)".replace(",", " ")]); ws["A1"].font = Font(bold=True, size=12)
ws.append(["Партия/№", "Позиция ТЗ", "Кол-во", "НМЦ ед., ₽", "Наша цена ед., ₽", "№ записи ПП 719", "Производитель (держатель, ИНН)", "Продукция в записи ГИСП", "Действует до", "Итог", "Комментарий"])
for c in ws[2]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
n_ok = n_eq = n_no = 0
for p in POS:
    o = OUT.get(p["tt"])
    st = "НЕТ ЗАПИСИ" if not o else "ЭКВИВАЛЕНТ" if o["note"].startswith("эквивалент") else "ПОДТВЕРЖДЕНО"
    n_ok += st == "ПОДТВЕРЖДЕНО"; n_eq += st == "ЭКВИВАЛЕНТ"; n_no += st == "НЕТ ЗАПИСИ"
    ws.append([p["tt"], p["name"], p["qty"], p["nmc"], p["price"], o and o["num"], o and o["man"], o and o["prod"], o and o["valid"], st, (o or {}).get("note") or ("ЗИП/комплектующие/импорт — российского изготовителя с записью ПП 719 не найдено; графы 6–8 оставлены «…»" if not o else "")])
    X = ws.max_row; ws.cell(X, 10).fill = PatternFill("solid", fgColor={"ПОДТВЕРЖДЕНО": "C6EFCE", "ЭКВИВАЛЕНТ": "DDEBF7", "НЕТ ЗАПИСИ": "F8CBAD"}[st])
    for j in (4, 5): ws.cell(X, j).number_format = "#,##0.00"
    for c in ws[X]: c.border = BOX; c.alignment = Alignment(wrap_text=True, vertical="top")
ws.append([]); ws.append([f"Подтверждено: {n_ok}; эквивалент с записью: {n_eq}; без записи: {n_no}"])
for i, w in enumerate([8, 50, 7, 12, 12, 11, 34, 46, 11, 14, 50], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "C3"; tidy(wb, row_h=36); wb.save(f"{OD}/02b_Реестр_сверка.xlsx")
print(f"позиций {len(POS)}; подтверждено {n_ok}, эквивалент {n_eq}, без записи {n_no}; итого {total:,.2f}")
