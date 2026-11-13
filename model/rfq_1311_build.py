"""Пакет запросов КП 13.11.2026: ссылки на закупки + спецификации + заводы с проверенными контактами + тексты писем.
→ ../ZAPROSY_KP_1311.xlsx (сводный) и data/rfq_1311/Spec_N.xlsx (вложения к письмам)"""
import json, re, os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

R = "data/rfq_1311/"; C = "data/companies/"
SPECS = json.load(open(R + "specs.json"))
def lj(f):
    try: d = json.load(open(C + f))
    except FileNotFoundError: return []
    return d if isinstance(d, list) else next((v for v in d.values() if isinstance(v, list)), [])
CAB, TUB, VLA = lj("rfq_cable_plants.json"), lj("rfq_tube_makers.json"), lj("rfq_video_light_appl.json")
def grp(x, *keys):
    g = (str(x.get("group", "")) + " " + " ".join(map(str, x.get("products", [])))).lower()
    return any(k in g for k in keys)
G = lambda k: [x for x in VLA if str(x.get("group", "")).endswith(k)]
FACT = {1: CAB, 2: TUB, 3: G("video"), 4: G("light"), 5: G("washers")}

H = PatternFill("solid", fgColor="0B2545"); HF = Font(bold=True, color="FFFFFF"); W = Alignment(wrap_text=True, vertical="top")
LINK = Font(color="0563C1", underline="single")
def head(ws, cols, widths):
    ws.append(cols)
    for c in ws[1]: c.fill = H; c.font = HF; c.alignment = W
    for i, w in enumerate(widths): ws.column_dimensions[chr(65 + i)].width = w
    ws.freeze_panes = "A2"

def url(num, etp):
    n = str(num)
    if re.fullmatch(r"\d{7}", n) and "b2b" in str(etp).lower(): return f"https://www.b2b-center.ru/market/view.html?id={n}"
    if n.startswith("z") and "/" in n: n = n.split("/")[1]
    if re.fullmatch(r"\d{11}|\d{19}", n): return f"https://zakupki.gov.ru/epz/order/extendedsearch/results.html?searchString={n}"
    if n == "17606-ТМЦ": return "https://tender.samolet.ru/"
    return None

wb = Workbook()
# 1. Ссылки
ws = wb.active; ws.title = "1. Закупки и ссылки"
head(ws, ["№", "Статус", "Номер", "Предмет", "Заказчик / регион", "Площадка", "НМЦ/оценка, ₽", "Подача", "Срок, мес", "Ссылка", "Спецификация №", "Что сделать"],
     [4, 18, 20, 44, 34, 16, 14, 12, 8, 40, 10, 50])
tab = load_workbook("../ZAKUPKI_1311_PODACHA_I_TZ.xlsx").active
spec_of = {l["lot"]: s["no"] for s in SPECS for l in s["lots"]}
for i, r in enumerate(tab.iter_rows(min_row=2, values_only=True), 1):
    num = str(r[3]); u = url(num, r[6])
    key = num.split("/")[-1] if "/" in num else num
    ws.append([i, r[2], num, r[4], r[5], r[6], r[7], r[8], r[9], u or "—", spec_of.get(num) or spec_of.get(key) or "", r[12]])
    if u: c = ws.cell(ws.max_row, 10); c.hyperlink = u; c.font = LINK
for row in ws.iter_rows(min_row=2):
    for c in row: c.alignment = W

# 2. Спецификации (сводно) + отдельные файлы
ws2 = wb.create_sheet("2. Спецификации")
head(ws2, ["Спец. №", "Группа", "Закупка", "Заказчик", "Подача", "№ п/п", "Наименование", "Кол-во", "Ед.", "Требования / характеристики"], [8, 30, 20, 30, 12, 6, 50, 10, 7, 80])
os.makedirs(R, exist_ok=True)
for s in SPECS:
    sw = Workbook(); sh = sw.active; sh.title = f"Спецификация {s['no']}"
    sh.append([f"Спецификация № {s['no']}: {s['title']}"]); sh["A1"].font = Font(bold=True)
    sh.append(["Просим указать: цену за ед. с НДС, наличие/срок поставки, страну производства, № реестровой записи (ГИСП / ПП 719) при наличии, условия оплаты, срок действия цены."]); sh.append([])
    sh.append(["№", "Закупка / объект", "Наименование", "Кол-во", "Ед.", "Требования", "Цена за ед. с НДС, ₽", "Наличие / срок", "Производитель, страна", "№ реестровой записи"])
    for c in sh[4]: c.fill = H; c.font = HF; c.alignment = W
    k = 0
    for l in s["lots"]:
        for j, p in enumerate(l["positions"], 1):
            k += 1
            ws2.append([s["no"], s["title"], l["lot"], l["customer"], l.get("deadline", ""), j, p.get("name"), p.get("qty"), p.get("unit"), p.get("req")])
            sh.append([k, f"{l['lot']} · {l['customer'][:40]}", p.get("name"), p.get("qty"), p.get("unit"), (p.get("req") or "")[:500], "", "", "", ""])
    for col, w in zip("ABCDEFGHIJ", [5, 34, 50, 10, 7, 70, 14, 14, 20, 16]): sh.column_dimensions[col].width = w
    for row in sh.iter_rows(min_row=4):
        for c in row: c.alignment = W
    sw.save(f"{R}Spec_{s['no']}.xlsx")
for row in ws2.iter_rows(min_row=2):
    for c in row: c.alignment = W

# 3. Заводы
ws3 = wb.create_sheet("3. Заводы и контакты")
head(ws3, ["Спец. №", "Компания", "ИНН", "Город", "Продукция", "E-mail", "Телефон", "Сайт", "Проверено (2 источника)", "Источники e-mail", "Источники телефона", "Примечание"],
     [8, 34, 13, 16, 40, 30, 20, 26, 12, 50, 50, 40])
for no, lst in FACT.items():
    for x in sorted(lst, key=lambda x: not x.get("verified")):
        ws3.append([no, x.get("company"), x.get("inn"), x.get("city"), ", ".join(map(str, x.get("products", [])))[:200], x.get("email"), x.get("phone"), x.get("site"),
                    "ДА" if x.get("verified") else "НЕТ", "\n".join(x.get("sources_email", []) or [])[:400], "\n".join(x.get("sources_phone", []) or [])[:400], (x.get("notes") or "")[:300]])
        if not x.get("verified"): ws3.cell(ws3.max_row, 9).font = Font(bold=True, color="B00020")
for row in ws3.iter_rows(min_row=2):
    for c in row: c.alignment = W

# 4. Письма
SIGN = "С уважением,\nАлександр Орлов\nООО «Гидропас»\ninfo@gidropas.ru"
def letter(s):
    lots = "; ".join(f"{l['customer'].split(',')[0][:40]} (подача {l.get('deadline') or 'см. таблицу'})" for l in s["lots"] if l["positions"])
    extra = {1: "Аналоги по ГОСТ допускаются. Для лотов с исполнением ХЛ укажите, пожалуйста, исполнение явно.",
             2: "Закупки по 44-ФЗ с запретом иностранных изделий (ПП 1875), поэтому нужен номер реестровой записи (ГИСП / ПП 719) и РУ. Если в ТЗ есть винтовая крышка или особая этикетка, напишите, можете ли вы это дать.",
             3: "Приоритет — оборудование из реестра российской продукции (ПП 1875, ограничение). Укажите номер реестровой записи.",
             4: "Допускаются аналоги, но размеры 1600×1600×80 и 161×73×170 должны совпадать без отклонений.",
             5: "Нужна официальная гарантия производителя на двигатель 120 месяцев и номер в реестре российской продукции."}[s["no"]]
    return (f"Тема: Запрос КП — {s['title'][:80]} (спецификация № {s['no']})\n\nДобрый день!\n\n"
            f"Мы участвуем в закупках: {lots}. Просим коммерческое предложение по приложенной спецификации № {s['no']}.\n\n"
            f"Укажите, пожалуйста: цену за единицу с НДС, наличие или срок изготовления, условия оплаты, доставку до объекта или самовывоз, срок действия цены.\n{extra}\n\n"
            f"Сроки подачи заявок короткие, будем благодарны за ответ до 15:00 ближайшего рабочего дня.\n\n{SIGN}")
ws4 = wb.create_sheet("4. Письма")
head(ws4, ["Спец. №", "Кому (e-mail, только проверенные)", "Вложение", "Текст письма"], [8, 50, 18, 120])
for s in SPECS:
    to = "; ".join(x["email"] for x in FACT.get(s["no"], []) if x.get("verified") and x.get("email") and not x.get("exclude"))
    ws4.append([s["no"], to or "— ждём проверки контактов —", f"Spec_{s['no']}.xlsx", letter(s)])
for row in ws4.iter_rows(min_row=2):
    for c in row: c.alignment = W
wb.save("../ZAPROSY_KP_1311.xlsx")
print({no: (len(l), sum(1 for x in l if x.get("verified"))) for no, l in FACT.items()})
