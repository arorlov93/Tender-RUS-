"""АО «ЭнергосбыТ Плюс» 32616535058 — хозяйственно-бытовые товары (запрос оферт, ЕЭТП, подача до 10.11 14:00): все 207 позиций из извещения ЕИС
(наименование/характеристика, ОКПД2, ед., цена за единицу заказчика, нацрежим). → ../ESPLUS_HOZ_32616535058.xlsx"""
import json, collections, openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from tidy import tidy
d = json.load(open("data/zakupki_0211/cards/32616535058.json"))
src = [x["source"] for x in d["docs"] if "lots" in x["source"]][-1]
lot = src["lots"]["lot"]; lot = lot if isinstance(lot, dict) else lot[0]
it = lot["lotData"]["lotItems"]["lotItem"]; it = it if isinstance(it, list) else [it]
thin = Side(style="thin", color="BFBFBF"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Позиции (207)"
ws.append(["АО «ЭнергосбыТ Плюс» № 32616535058 — поставка хозяйственно-бытовых товаров. НМЦ договора 12 275 780,47 ₽; запрос оферт на ЕЭТП (com.roseltorg.ru, процедура 1047553); подача до 10.11.2026 14:00, итоги 24.11.2026"]); ws["A1"].font = Font(bold=True, size=12)
ws.append(["Контакт: Малькова Ю. Н., +7 (3532) 34-74-99, Yuliya.Malkova@esplus.ru. Количество в извещении — 1 по каждой позиции (цены за единицу); объёмы и адреса — в документации (zip)."])
ws.append(["№", "Товар (характеристика из извещения)", "ОКПД2", "Наименование по ОКПД2", "Ед.", "Кол-во в извещении", "Цена за ед. заказчика, ₽", "Нацрежим"])
for c in ws[3]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
G = collections.defaultdict(lambda: [0, 0.0]); nr = collections.Counter()
for x in it:
    reg = "запрет" if x.get("prohibitionNationalRegime") == "true" else "ограничение" if x.get("constraintNationalRegime") == "true" else "исключение" if x.get("exceptionNationalRegime") == "true" else "—"
    p = float(x.get("commodityItemPrice") or 0)
    ws.append([int(x["ordinalNumber"]), x.get("additionalInfo"), x["okpd2"]["code"], x["okpd2"]["name"], x["okei"]["name"], float(x["qty"]), p, reg])
    X = ws.max_row; ws.cell(X, 7).number_format = "#,##0.00"
    for c in ws[X]: c.border = BOX; c.alignment = Alignment(wrap_text=True, vertical="top")
    g = G[x["okpd2"]["name"]]; g[0] += 1; g[1] += p; nr[reg] += 1
for i, w in enumerate([5, 60, 14, 40, 10, 9, 14, 13], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "C4"; ws.auto_filter.ref = f"A3:H{ws.max_row}"
s2 = wb.create_sheet("По группам ОКПД2")
s2.append(["Группа (ОКПД2)", "Позиций", "Сумма цен за ед., ₽"])
for c in s2[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")
for k, v in sorted(G.items(), key=lambda x: -x[1][0]): s2.append([k, v[0], round(v[1], 2)])
s2.column_dimensions["A"].width = 70; s2.column_dimensions["C"].width = 18
tidy(wb, row_h=30); wb.save("../ESPLUS_HOZ_32616535058.xlsx")
print(len(it), "позиций; нацрежим:", dict(nr), "; сумма цен ед.:", round(sum(float(x.get("commodityItemPrice") or 0) for x in it), 2))
for k, v in sorted(G.items(), key=lambda x: -x[1][0])[:12]: print(f"  {v[0]:3} {k[:70]}")
