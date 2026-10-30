"""Сахалинэнерго 11201/11202: разбор КП от 28.10.2026 (ПЭК, Лидер-Энерго, ВФЗ, Доминик, АРМАТЭК) — цена с НДС против НМЦ без НДС (мы на УСН) и прежней оценки. → ../SAKHALIN_KP_2810.xlsx"""
import json, openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side
from tidy import tidy
K = json.load(open("data/sakh/kp/kp_2810.json")); P = json.load(open("data/sakh/nmc_positions.json"))
OLD = {"ПЭК 11201/1/67": 345046.5, "ПЭК 11202/1/33": 55258.2, "ПЭК 11202/1/64": 64287.9}  # прежняя оценка (до КП) — для сравнения
thin = Side(style="thin", color="BFBFBF"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin); M = '#,##0" ₽"'
head = lambda ws: [setattr(c, "font", Font(bold=True, color="FFFFFF")) or setattr(c, "fill", PatternFill("solid", fgColor="305496")) for c in ws[ws.max_row] if c.value is not None]
wb = openpyxl.Workbook(); sv = wb.active; sv.title = "Итог по поставщикам"
sv.append(["Сахалинэнерго 11201 / 11202 — КП от 28.10.2026: что предлагают и выгодно ли (наша цена = НМЦ без НДС, закупка с НДС; доставка до Южно-Сахалинска — отдельно)"]); sv["A1"].font = Font(bold=True, size=12)
sv.append(["Поставщик", "Документ", "Позиций", "Сумма КП с НДС, ₽", "НМЦ этих позиций, ₽", "Маржа", "Условия", "Вывод"]); head(sv)
rows = []; agg = {}
for x in K["items"]:
    p = next(p for p in P[x["lot"]] if p["n"] == x["n"] and p["part"].endswith(str(x["part"])))
    v = K["suppliers"][x["s"]]["vat"]; pv = x["price"] * (1.22 if v is False else 1); m = 1 - pv / p["nmc_unit"]
    verdict = "выгодно" if m >= 0.2 else "торг / тонко" if m >= 0.05 else "не брать — дороже НМЦ" if m < 0 else "на грани (без доставки)"
    rows.append([x["s"], K["suppliers"][x["s"]]["doc"], x["lot"], p["part"], x["n"], p["name"], x["name"], x["q"], x["price"], "без НДС" if v is False else ("с НДС" if v else "не указан"),
                 pv, p["nmc_unit"], m, pv * x["q"], p["nmc_unit"] * x["q"], verdict, x["note"]])
    a = agg.setdefault(x["s"], [0, 0, 0]); a[0] += 1; a[1] += pv * x["q"]; a[2] += p["nmc_unit"] * x["q"]
CON = {"ПЭК": "№1 (МЭОФ-1600 ЕА) выгодно 43%; №2 МЭОФ-1600 96К — 9%; №3 МЭОФ-250 99К — 3% и исполнение не 97К как в ТЗ → торг/уточнить",
       "Лидер-Энерго": "выгодно 38%; подтвердить полное обозначение ЗДП-80-6,3-1-М-А-УХЛ1", "ВФЗ": "выгодно 28–51%; Ду32 Ру16 вместо Ру10 — согласовать",
       "Доминик": "выгодно 37–58%; уточнить НДС и изготовителя Т-229Б (реестр ПП 719)", "rost (ФОБОС ФБ39)": "выгодно 55%: 30,3 тыс. с НДС при НМЦ 66,8 тыс.; Универсал из нержавейки вместо Классик Ст20 — эквивалент, согласовать; реквизиты и № реестра запросить", "ЛМЗ": "30ч6бр Ду50–250 для лота 2: маржа 21–38% даже с КОФ (ВФЗ); позиции лота 1 в счёте не нужны; ЛМЗ — окр. среда от −15 °С: сверить с климатом ТЗ", "Ярдос": "6 позиций ЗАРД/ЗАРДП — точно по ТЗ, маржа 14–20%; № реестра запросить", "Гималаи": "КШ250.8Д1 ×6 — точно по ТЗ, маржа 16%, доставка до Ю-Сахалинска включена", "ТД АДЛ": "лот 2: ЗПВС Ду100/150 — 15%; Гранлок шаровой DN150 — 6%, DN100 — минус 2% (торг); остальные 12 строк КП — лот 1, не нужны", "АРМАТЭК": "в целом 6,5% — тонко; поз. 38 (АН6 DN065 пневмо) дороже НМЦ на 40% → минус 180 тыс.; ЗИП 12–16%; КП действует 5 банк. дней!"}
for s, a in agg.items():
    sv.append([s, K["suppliers"][s]["doc"], a[0], a[1], a[2], 1 - a[1] / a[2], K["suppliers"][s]["terms"], CON[s]])
    X = sv.max_row; sv.cell(X, 4).number_format = M; sv.cell(X, 5).number_format = M; sv.cell(X, 6).number_format = "0%"
    for c in sv[X]: c.border = BOX
t1, t2 = sum(a[1] for a in agg.values()), sum(a[2] for a in agg.values())
sv.append(["ИТОГО", "", sum(a[0] for a in agg.values()), t1, t2, 1 - t1 / t2]); sv.cell(sv.max_row, 4).number_format = M; sv.cell(sv.max_row, 5).number_format = M; sv.cell(sv.max_row, 6).number_format = "0%"
for c in sv[sv.max_row]: c.font = Font(bold=True)
sv.append([]); sv.append(["Влияние на лоты (снижение 10%, УСН): лот 2 (11202) прибыль ≈ 2,84 млн (было 3,06; ЛМЗ 30ч6бр дороже прежней оценки на 0,17 млн — МЭОФ-1600 96К по КП 172 тыс. против оценки 55 тыс.); лот 1 (11201) ≈ −0,38 млн (АРМАТЭК поз. 38)."])
for i, w in enumerate([16, 26, 9, 16, 16, 8, 60, 70], 1): sv.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws = wb.create_sheet("Построчно")
ws.append(["Поставщик", "Документ", "Лот", "Партия", "№", "Позиция ТЗ", "Предложено", "Кол-во", "Цена в КП, ₽", "НДС в КП", "Цена с НДС, ₽", "НМЦ за ед. без НДС, ₽", "Маржа", "Сумма с НДС, ₽", "НМЦ сумма, ₽", "Вывод", "Что уточнить"]); head(ws)
COL = {"выгодно": "C6EFCE", "торг / тонко": "FFEB9C", "на грани (без доставки)": "FCE4D6", "не брать — дороже НМЦ": "F8CBAD"}
for r in rows:
    ws.append(r); X = ws.max_row
    for j in (9, 11, 12, 14, 15): ws.cell(X, j).number_format = M
    ws.cell(X, 13).number_format = "0%"; ws.cell(X, 16).fill = PatternFill("solid", fgColor=COL[r[15]])
    for c in ws[X]: c.border = BOX
for i, w in enumerate([12, 22, 7, 12, 5, 40, 40, 7, 12, 9, 12, 12, 8, 13, 13, 16, 40], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "G2"; ws.auto_filter.ref = f"A1:Q{ws.max_row}"
tidy(wb, row_h=40); wb.save("../SAKHALIN_KP_2810.xlsx"); print("ok", round(t1), round(t2))
