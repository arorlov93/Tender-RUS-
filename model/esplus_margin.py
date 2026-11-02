"""ЭнергосбыТ Плюс 32616535058: маржа по рыночным ценам на 207 позиций хозтоваров.
Цена заказчика — за единицу из извещения ЕИС; рынок — цены российских магазинов/оптовиков с НДС (data/zakupki_0211/esplus/out*.json).
Маржа = (цена заказчика − рынок) / цена заказчика; при снижении −10 % на торгах. Доставка не учтена (адреса — в документации).
→ ../ESPLUS_HOZ_MARGIN.xlsx"""
import json, glob, openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from tidy import tidy

P = {x["n"]: x for f in sorted(glob.glob("data/zakupki_0211/esplus/part*.json")) for x in json.load(open(f))}
O = {}
for f in sorted(glob.glob("data/zakupki_0211/esplus/out*.json")):
    for x in json.load(open(f)): O[int(x["n"])] = x
thin = Side(style="thin", color="BFBFBF"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin); M = "#,##0.00"
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Маржа по позициям"
ws.append(["АО «ЭнергосбыТ Плюс» 32616535058 — хозтовары: цена заказчика за ед. против рыночной закупочной цены (с НДС). Без доставки."]); ws["A1"].font = Font(bold=True, size=12)
ws.append(["№", "Товар (ТЗ)", "Ед.", "Цена заказчика, ₽", "Рынок, ₽", "Маржа", "Маржа при −10 %", "Совпадение", "Цена за что", "Источник", "Ссылка"])
for c in ws[2]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
S = dict(n=0, found=0, nmc=0.0, mkt=0.0, nmc_f=0.0, neg=0, hi=0)
for n in sorted(P):
    p = P[n]; o = O.get(n, {}); mk = o.get("price"); nmc = float(p["nmc"])
    m = (nmc - mk) / nmc if mk and nmc else None; m10 = (nmc * 0.9 - mk) / (nmc * 0.9) if mk and nmc else None
    ws.append([n, p["name"], p["unit"], nmc, mk, m, m10, o.get("match"), o.get("unit_note"), o.get("source"), o.get("url")])
    X = ws.max_row
    for j in (4, 5): ws.cell(X, j).number_format = M
    for j in (6, 7): ws.cell(X, j).number_format = "0%"
    if m10 is not None: ws.cell(X, 7).fill = PatternFill("solid", fgColor="C6EFCE" if m10 >= 0.2 else "FFF2CC" if m10 >= 0 else "F8CBAD")
    if o.get("url"): ws.cell(X, 11).hyperlink = o["url"]
    for c in ws[X]: c.border = BOX; c.alignment = Alignment(wrap_text=True, vertical="top")
    S["n"] += 1; S["nmc"] += nmc
    if mk: S["found"] += 1; S["nmc_f"] += nmc; S["mkt"] += mk; S["neg"] += m10 < 0; S["hi"] += m10 >= 0.2
for i, w in enumerate([5, 55, 9, 12, 12, 8, 9, 22, 28, 18, 40], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "C3"; ws.auto_filter.ref = f"A2:K{ws.max_row}"
s2 = wb.create_sheet("Итог", 0)
mg = (S["nmc_f"] - S["mkt"]) / S["nmc_f"] if S["nmc_f"] else 0; mg10 = (S["nmc_f"] * 0.9 - S["mkt"]) / (S["nmc_f"] * 0.9) if S["nmc_f"] else 0
for r in [["Итог по корзине (по 1 ед. каждой позиции)"], ["Позиций всего", S["n"]], ["Найдена рыночная цена", S["found"]],
          ["Сумма цен заказчика по найденным, ₽", round(S["nmc_f"], 2)], ["Сумма рыночных цен по найденным, ₽", round(S["mkt"], 2)],
          ["Средняя маржа (по корзине)", mg], ["Средняя маржа при −10 %", mg10], ["Позиций с маржой ≥ 20 % при −10 %", S["hi"]], ["Позиций в убытке при −10 %", S["neg"]],
          ["НМЦ договора 12 275 780,47 ₽ — лимит; объёмы по позициям в документации (zip). Маржа по корзине — ориентир, итог зависит от структуры объёмов."]]:
    s2.append(r)
s2["A1"].font = Font(bold=True, size=12)
for r in (6, 7): s2.cell(r, 2).number_format = "0%"
for r in (4, 5): s2.cell(r, 2).number_format = M
s2.column_dimensions["A"].width = 48; s2.column_dimensions["B"].width = 18
tidy(wb, row_h=30); wb.save("../ESPLUS_HOZ_MARGIN.xlsx")
print(S, f"маржа {mg:.1%}, при −10% {mg10:.1%}")
