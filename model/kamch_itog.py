"""Камчатка 32616346943 — ИТОГ на 29.10.2026: себестоимость лота по полученным КП и нашим ценам, прибыль по сценариям снижения, что осталось получить. → ../KAMCHATKA_ITOG_2910.xlsx"""
import json, runpy, io, contextlib, collections, openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side
from tidy import tidy

with contextlib.redirect_stdout(io.StringIO()):
    rows = runpy.run_path("kamch_status.py")["rows"]
    FIN = runpy.run_path("kamch_final.py")["SEL"]
REG = json.load(open("data/kamch/registry.json"))
N, FREIGHT, TORG = 101102908, 3069024, 0.10
KPA = ("Брать", "Брать при условии", "Торг: наш бюджет")
for d in rows:
    d["cost"] = d["inv"] if d["act"] in KPA else d["our"] * (1 - TORG)
    d["src"] = f'КП {d["sup"]}' if d["act"] in KPA else "наша проверенная цена −10% (КП нет)"
C = sum(d["cost"] * d["q"] for d in rows) + FREIGHT
def net(dsc):
    P = N * (1 - dsc); pre = P - C - 0.008 * P - C * 0.2 * 120 / 365; return pre - 0.15 * max(pre, 0)
thin = Side(style="thin", color="BFBFBF"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin); M = '#,##0" ₽"'
head = lambda ws: [setattr(c, "font", Font(bold=True, color="FFFFFF")) or setattr(c, "fill", PatternFill("solid", fgColor="305496")) for c in ws[ws.max_row] if c.value is not None]
def box(ws, r0):
    for row in ws.iter_rows(min_row=r0):
        for c in row:
            if c.value is not None: c.border = BOX
wb = openpyxl.Workbook(); s = wb.active; s.title = "Итог"
s.append(["Камчатскэнерго 32616346943 — итог на 29.10.2026 (подача 02.11.2026 10:00 МСК)"]); s["A1"].font = Font(bold=True, size=13)
s.append(["Себестоимость: по полученным КП (брать / при условии / торг — по цене КП) + наша проверенная цена −10% там, где КП нет; доставка до Петропавловска-Камчатского 55 ₽/кг. Цены с НДС (мы на УСН)."])
s.append([]); s.append(["Показатель", "Значение"]); head(s); r0 = s.max_row
for a, b in [("НМЦК", N), ("Товар по КП и нашим ценам", C - FREIGHT), ("Доставка", FREIGHT), ("Себестоимость лота", C), ("Маржа при НМЦК", 1 - C / N),
             ("Прежний расчёт (до КП, 20.10)", 54810725), ("Изменение себестоимости", C - 54810725)]:
    s.append([a, b]); s.cell(s.max_row, 2).number_format = "0.0%" if "Маржа" in a else M
box(s, r0); s.append([])
CB = sum((d["inv"] if d["inv"] is not None else d["our"] * (1 - TORG)) * d["q"] for d in rows) + FREIGHT  # сценарий Б: дорогие позиции берём по КП
def netc(Cx, dsc):
    P = N * (1 - dsc); pre = P - Cx - 0.008 * P - Cx * 0.2 * 120 / 365; return pre - 0.15 * max(pre, 0)
s.append(["Снижение в заявке", "Цена заявки", "А: торг (31 поз.) по КП, 29 дорогих — на заводы по нашей цене", "Б: всё по КП, включая 29 поз. «заказать у других» по цене ДН.РУ/Тулы"]); head(s); r0 = s.max_row
s.append(["Себестоимость лота", None, C, CB]); s.cell(s.max_row, 3).number_format = M; s.cell(s.max_row, 4).number_format = M
s.append(["Маржа при НМЦК", None, 1 - C / N, 1 - CB / N]); s.cell(s.max_row, 3).number_format = "0.0%"; s.cell(s.max_row, 4).number_format = "0.0%"
for dsc in (0, 0.1, 0.2, 0.3, 0.4):
    s.append([dsc, N * (1 - dsc), netc(C, dsc), netc(CB, dsc)]); X = s.max_row; s.cell(X, 1).number_format = "0%"
    for j in (2, 3, 4): s.cell(X, j).number_format = M
box(s, r0); s.append([])
s.append(["Статус позиции", "Позиций", "Сумма в себестоимости", "Наша цена", "Цена по КП", "Источник"]); head(s); r0 = s.max_row
for a in ("Брать", "Брать при условии", "Торг: наш бюджет", "Заказать у других", "Не отработано"):
    g = [d for d in rows if d["act"] == a]
    s.append([a, len(g), sum(d["cost"] * d["q"] for d in g), sum(d["our"] * d["q"] for d in g), sum((d["inv"] or 0) * d["q"] for d in g) or None, "КП" if a in KPA else "наша цена −10%"])
    for j in (3, 4, 5): s.cell(s.max_row, j).number_format = M
s.append(["ИТОГО", len(rows), C - FREIGHT, sum(d["our"] * d["q"] for d in rows)]); [setattr(c, "font", Font(bold=True)) for c in s[s.max_row]]
for j in (3, 4): s.cell(s.max_row, j).number_format = M
box(s, r0)
for i, w in enumerate([46, 18, 26, 30, 18, 22], 1): s.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

o = wb.create_sheet("Что осталось")
o.append(["Что осталось получить до подачи", "Позиций", "Сумма, ₽", "Что сделать"]); head(o)
need = [d for d in rows if d["act"] in ("Заказать у других", "Не отработано")]
nonum = [d for d in rows if not REG.get(d["tt"], {}).get("num")]
dn = [d for d in rows if d["sup"] == "ДН.РУ" and d["act"] in ("Брать", "Брать при условии") and not REG.get(d["tt"], {}).get("num")]
nz = [t for t, v in FIN.items() if v["src"] == "НЕ ЗАКРЫТО"]
for a, g, w in [("КП от заводов (нет цены / ДН.РУ дорого или не по ТЗ)", need, "KAMCHATKA_FINAL_ZAVODY — 22 завода; ответы до 31.10"),
                ("Торг с ДН.РУ (КП дороже нашей цены на 10–30%)", [d for d in rows if d["act"] == "Торг: наш бюджет"], "назвать бюджет (наша цена −10%)"),
                ("Подтверждения «при условии» (реестр / эквивалент / КОФ / CF8)", [d for d in rows if d["act"] == "Брать при условии"], "письма поставщикам — KAMCHATKA_KP_VYGODNO"),
                ("Без № реестровой записи ПП 719", nonum, "выписки ГИСП от заводов; иначе — к иностранной продукции"),
                ("из них ДН.РУ — выгодные позиции без номера", dn, "если ДН.РУ не даст изготовителя с записью — перевести на заводы"),
                ("Не закрыто совсем (нет рос. производителя, проходящего по ТЗ)", [d for d in rows if d["tt"] in nz], "запрос разъяснений заказчику (чугун→сталь, температура, ВДХ, КШХ)")]:
    o.append([a, len(g), sum(d["our"] * d["q"] for d in g), w]); o.cell(o.max_row, 3).number_format = M
box(o, 1)
for i, w in enumerate([58, 10, 18, 70], 1): o.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

p = wb.create_sheet("По поставщикам КП")
p.append(["Поставщик", "Позиций в КП", "Из них берём", "Сумма КП (берём), ₽", "Наша цена этих позиций, ₽", "Разница"]); head(p)
for sup in sorted({d["sup"] for d in rows if d["sup"]}):
    g = [d for d in rows if d["sup"] == sup]; t = [d for d in g if d["act"] in KPA]
    a, b = sum(d["inv"] * d["q"] for d in t), sum(d["our"] * d["q"] for d in t)
    p.append([sup, len(g), len(t), a, b, (a / b - 1) if b else None]); X = p.max_row
    p.cell(X, 4).number_format = M; p.cell(X, 5).number_format = M; p.cell(X, 6).number_format = "+0%;-0%"
box(p, 1)
for i, w in enumerate([30, 12, 12, 18, 20, 10], 1): p.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

e = wb.create_sheet("Дороже бюджета (КП есть)")
e.append(["Табл./№", "Позиция ТЗ", "Кол-во", "Поставщик", "Что в счёте", "Цена КП, ₽/шт", "Наша цена, ₽/шт", "Переплата, ₽", "Отклонение", "Почему", "Альтернатива (завод с реестром)"]); head(e)
for d in sorted([d for d in rows if d["act"] == "Заказать у других"], key=lambda d: -(d["inv"] - d["our"]) * d["q"]):
    f = FIN.get(d["tt"], {})
    e.append([d["tt"], d["name"], d["q"], d["sup"], d["offer"][:110], d["inv"], d["our"], (d["inv"] - d["our"] * (1 - TORG)) * d["q"], d["inv"] / d["our"] - 1, d["why"][:120], f.get("who", "")])
    X = e.max_row
    for j in (6, 7, 8): e.cell(X, j).number_format = M
    e.cell(X, 9).number_format = "+0%;-0%"
gx = [d for d in rows if d["act"] == "Заказать у других"]
e.append(["", "ИТОГО", sum(d["q"] for d in gx), "", "", sum(d["inv"] * d["q"] for d in gx), sum(d["our"] * d["q"] for d in gx), sum((d["inv"] - d["our"] * (1 - TORG)) * d["q"] for d in gx)])
for j in (6, 7, 8): e.cell(e.max_row, j).number_format = M
for c in e[e.max_row]: c.font = Font(bold=True)
box(e, 1)
for i, w in enumerate([8, 36, 7, 18, 40, 14, 14, 14, 10, 30, 34], 1): e.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
a = wb.create_sheet("Все 315 позиций")
a.append(["Табл./№", "Позиция ТЗ", "Кол-во", "Наша цена, ₽/шт", "Цена КП, ₽/шт", "В себестоимости, ₽/шт", "Сумма, ₽", "Источник", "Статус", "Поставщик (выбор)", "№ реестра ПП 719", "Проверка ТЗ / что уточнить"]); head(a)
COL = {"Брать": "C6EFCE", "Брать при условии": "E2EFDA", "Торг: наш бюджет": "FFEB9C", "Заказать у других": "F8CBAD", "Не отработано": "DDEBF7"}
for d in sorted(rows, key=lambda d: (d["tt"].split("/")[0], int(d["tt"].split("/")[1]))):
    f = FIN.get(d["tt"], {})
    a.append([d["tt"], d["name"], d["q"], d["our"], d["inv"], d["cost"], d["cost"] * d["q"], d["src"], d["act"], f.get("who") or d["sup"], REG.get(d["tt"], {}).get("num") or "",
              ((f.get("verdict") or "") + ("; " + f["ask"] if f.get("ask") else ""))[:250]])
    X = a.max_row
    for j in (4, 5, 6, 7): a.cell(X, j).number_format = M
    a.cell(X, 9).fill = PatternFill("solid", fgColor=COL[d["act"]])
box(a, 1)
for i, w in enumerate([8, 44, 7, 13, 13, 13, 14, 26, 16, 30, 12, 50], 1): a.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
a.freeze_panes = "C2"; a.auto_filter.ref = f"A1:L{a.max_row}"
tidy(wb, row_h=30); wb.save("../KAMCHATKA_ITOG_2910.xlsx")
print(round(C), round(CB), round(1 - C / N, 3), round(net(0.1)), round(netc(CB, 0.1)), len(need), len(nonum), len(dn), len(nz))
