"""АНО «Таврида.Арт» B1510261835045 — форменная одежда СБ: себестоимость по ТЗ, маржа, цены для заявки."""
import json, openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side
from tidy import tidy

D = json.load(open("data/tavrida/prices.json"))
NMCK = 4169750
QTY = {1: 145, 2: 145, 3: 145, 4: 145, 5: 145, 6: 145, 7: 170, 8: 170}
NAMES = {1: "Тактический костюм (демисезонный)", 2: "Тактическая кофта (флисовая)", 3: "Тактическая куртка с капюшоном (зимняя)", 4: "Тактическая шапка (зимняя)",
         5: "Полуботинки (утепленные)", 6: "Жилет разгрузочный тактический", 7: "Тактические перчатки (зимние)", 8: "Тактический снуд"}
DELIV = D["delivery"]["cost"]
items = {x["n"]: x for x in D["items"]}
C = sum(items[n]["recommended_cost"] * QTY[n] for n in QTY) + DELIV


def model(P, C, days=60):
    pre = P - C - 0.008 * P - C * 0.2 * days / 365
    return pre - 0.15 * max(pre, 0)


def be():
    lo, hi = 0, 0.9
    for _ in range(60):
        m = (lo + hi) / 2
        lo, hi = (m, hi) if model(NMCK * (1 - m), C) > 0 else (lo, m)
    return lo


thin = Side(style="thin", color="BFBFBF"); M = '#,##0" ₽"'; P = "0.0%"
def head(ws):
    for c in ws[ws.max_row]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")
def box(ws, r0):
    for row in ws.iter_rows(min_row=r0):
        for c in row: c.border = Border(top=thin, bottom=thin, left=thin, right=thin)

wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Себестоимость"
ws.append(["АНО «Таврида.Арт» B1510261835045 — форменная одежда службы безопасности (Судак). Подача 30.10.2026 12:00 МСК, Росэлторг (коммерческий тендер)"]); ws["A1"].font = Font(bold=True, size=12)
ws.append(["По ТЗ (Прил. 1 к договору № 1193): поставка за 10 дней с даты договора, оплата 30 календарных дней после приёмки, аванса нет, гарантия 12 мес. Цены поставщиков 2026 с НДС."])
ws.append(["№", "Позиция", "Кол-во", "Себес за ед., ₽", "Сумма себеса, ₽", "Вес ед., кг", "Основной источник цены", "Подтверждение", "Ссылка"]); head(ws)
for n in sorted(QTY):
    it = items[n]; s = (it.get("sources") or [{}])[0]
    ws.append([n, NAMES[n], QTY[n], it["recommended_cost"], it["recommended_cost"] * QTY[n], it.get("weight_kg"), f'{s.get("company", "")}: {s.get("price", "")} ₽ — {s.get("note", "")}'[:150],
               s.get("date_evidence", "")[:40], s.get("url", "")])
    r = ws.max_row; ws.cell(r, 4).number_format = M; ws.cell(r, 5).number_format = M
    if s.get("url"): ws.cell(r, 9).hyperlink = s["url"]
ws.append(["", "Доставка Москва → Судак (~0,95 т), оценка", "", "", DELIV, "", D["delivery"].get("note", "")[:150]]); ws.cell(ws.max_row, 5).number_format = M
ws.append(["", "ИТОГО себестоимость", "", "", C]); ws.cell(ws.max_row, 5).number_format = M
for c in ws[ws.max_row]: c.font = Font(bold=True)
box(ws, 4)
for i, w in enumerate([4, 38, 8, 13, 15, 9, 60, 22, 45], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

m = wb.create_sheet("Маржа")
m.append(["Показатель", "Значение"]); head(m)
for a, b, f in [("НМЦК, ₽", NMCK, M), ("Себестоимость, ₽", C, M), ("Маржа при НМЦК", 1 - C / NMCK, P), ("Безубыточное снижение", be(), P)]:
    m.append([a, b]); m.cell(m.max_row, 2).number_format = f
m.append([]); m.append(["Снижение", "Цена заявки, ₽", "Прибыль после УСН, ₽", "Маржа к цене"]); head(m)
for d in (0, 0.05, 0.10, 0.15, 0.20, 0.25):
    Pp = NMCK * (1 - d); n_ = model(Pp, C); m.append([d, Pp, n_, n_ / Pp])
    r = m.max_row; m.cell(r, 1).number_format = P; m.cell(r, 2).number_format = M; m.cell(r, 3).number_format = M; m.cell(r, 4).number_format = P
box(m, 1)
for i, w in enumerate([24, 18, 20, 14], 1): m.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

# цены для заявки: −10% от НМЦК, распределены пропорционально себестоимости, целые рубли, итог сходится
BID = round(NMCK * 0.90)
units = {n: round(items[n]["recommended_cost"] * BID / (C - DELIV)) for n in QTY}
diff = BID - sum(units[n] * QTY[n] for n in QTY)
units[1] += round(diff / QTY[1])
sp = wb.create_sheet("Цены для заявки")
sp.append(["Спецификация (Прил. 2) — наше предложение, ≈ −10% от НМЦК; доставка включена в цены"]); sp["A1"].font = Font(bold=True)
sp.append(["№", "Перечень товаров", "Ед.", "Кол-во", "Цена за ед., руб.", "Сумма, руб."]); head(sp)
for n in sorted(QTY):
    sp.append([n, NAMES[n], "шт.", QTY[n], units[n], units[n] * QTY[n]])
    sp.cell(sp.max_row, 5).number_format = M; sp.cell(sp.max_row, 6).number_format = M
tot = sum(units[n] * QTY[n] for n in QTY)
sp.append(["", "ИТОГО (НДС не облагается, УСН)", "", "", "", tot]); sp.cell(sp.max_row, 6).number_format = M
for c in sp[sp.max_row]: c.font = Font(bold=True)
sp.append(["", f"Снижение от НМЦК: {1 - tot / NMCK:.1%}; прибыль после УСН ≈ {model(tot, C):,.0f} ₽".replace(",", " ")])
box(sp, 2)
for i, w in enumerate([4, 40, 6, 8, 16, 16], 1): sp.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
tidy(wb, row_h=34)
wb.save("../TAVRIDA_B1510261835045.xlsx")
print("C", C, "margin", 1 - C / NMCK, "be", be(), "bid", tot, "net@bid", round(model(tot, C)), units)
