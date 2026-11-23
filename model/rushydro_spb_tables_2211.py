"""Отдельные таблицы по двум закупкам, на которые идём: РусГидро (Нерюнгринская ГРЭС-2) и АО «Теплосеть СПб».
Позиции, НМЦ, себестоимость (оценка), цена в заявку с разбросом, маржа. → ../TABLICY_RUSGIDRO_TEPLOSET_2211.xlsx"""
import json, random
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
random.seed(2211)
r2 = lambda x: round(x + 1e-9, 2)
N = json.load(open("data/zakupki_2011_valves/neryungri/analysis.json"))
S = json.load(open("data/zakupki_2011_valves/spb/analysis.json"))

def fit(items, target, lo, hi, cap=9e9):
    k = [random.uniform(lo, hi) for _ in items]
    for _ in range(500):
        s = target / sum(b * q * ki for (b, q, f), ki in zip(items, k))
        k = [min(max(ki * s, f / b), cap) for (b, q, f), ki in zip(items, k)]
    pr = [r2(b * ki) for (b, q, f), ki in zip(items, k)]
    diff = r2(target - sum(p * q for p, (b, q, f) in zip(pr, items)))
    j = next((i for i, (b, q, f) in enumerate(items) if q == 1), None)
    if j is not None: pr[j] = r2(pr[j] + diff)
    return pr

wb = Workbook(); wb.remove(wb.active)
H = PatternFill("solid", fgColor="0B2545"); HF = Font(bold=True, color="FFFFFF"); Y = PatternFill("solid", fgColor="FFF2CC")
def sheet(title, head, rows, widths, money_cols):
    ws = wb.create_sheet(title); ws.append(head)
    for c in ws[1]: c.fill = H; c.font = HF; c.alignment = Alignment(wrap_text=True, vertical="top")
    for r in rows: ws.append(r)
    for i, w in enumerate(widths, 1): ws.column_dimensions[ws.cell(1, i).column_letter].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if c.column in money_cols and isinstance(c.value, (int, float)): c.number_format = "# ##0"
    ws.freeze_panes = "C2"; return ws

# ---------------- РусГидро / Нерюнгри: цены без НДС, скидка −20 % к НМЦ, разброс по позициям
P = N["positions"]; NMC_N = sum(p["nmc_unit"] * p["qty"] for p in P); TN = r2(NMC_N * 0.80)
over = N["economics"]["usn_base"]["overheads"]["total"]; goods = N["economics"]["usn_base"]["goods"]
cu = [p["cost_unit_base_vat"] * (1 + over / goods) for p in P]  # себестоимость с распределёнными накладными
pn = fit([(p["nmc_unit"], p["qty"], c * 1.10) for p, c in zip(P, cu)], TN, 0.70, 0.90, cap=0.98)
rows = []
for p, c, pr in zip(P, cu, pn):
    rows.append([p["n"], p["name"], p["qty"], p["nmc_unit"], r2(p["nmc_unit"] * p["qty"]), round(c), round(c * p["qty"]), pr, r2(pr * p["qty"]),
                 round((1 - pr / p["nmc_unit"]) * 100, 1), round(pr * p["qty"] - c * p["qty"]), round((1 - c / pr) * 100, 1), p.get("cost_basis", "")[:160]])
tc = sum(r[6] for r in rows); tb = sum(r[8] for r in rows)
rows.append(["", "ИТОГО", sum(p["qty"] for p in P), "", r2(NMC_N), "", round(tc), "", r2(tb), round((1 - tb / NMC_N) * 100, 1), round(tb - tc), round((1 - tc / tb) * 100, 1), ""])
ws = sheet("РусГидро Нерюнгри", ["№", "Позиция", "Кол-во", "НМЦ ед. без НДС", "НМЦ сумма", "Себестоимость ед. (оценка, с накладными)", "Себестоимость сумма",
           "ЦЕНА ЕД. В ЗАЯВКУ", "Сумма заявки", "Скидка к НМЦ, %", "Прибыль до налога", "Маржа, %", "Основа себестоимости"], rows, [4, 34, 7, 13, 14, 14, 14, 14, 14, 9, 14, 8, 50], {4, 5, 6, 7, 8, 9, 11})
for c in ws[ws.max_row]: c.font = Font(bold=True)
ws.append([]); ws.append(["", "Сценарии (весь лот, без НДС)", "", "Цена", "", "", "Себест. база", "", "Прибыль база", "", "Себест. пессим.", "Прибыль пессим."])
for lab, k in [("НМЦ", 0), ("−10 %", .1), ("−20 % (в таблице)", .2), ("−30 %", .3), ("−35 %", .35)]:
    pr = NMC_N * (1 - k); cb = N["economics"]["usn_base"]["total_cost"]; cp = N["economics"]["usn_pess"]["total_cost"]
    ws.append(["", lab, "", round(pr), "", "", round(cb), "", round(pr - cb), "", round(cp), round(pr - cp)])
ws.append(["", "Себестоимость — оценка без КП заводов (ПромАрм, ЗЭиМ, Тулаэлектропривод, САЗ «Авангард», Сталь-Контур). Накладные (доставка Нерюнгри 1,3 млн, шеф-монтаж 0,5, гарантия/финансирование 1,17, резерв 10 %) распределены пропорционально. Прибыль до налога; при переходе на НДС 22 % — с вычетом входного."])
ws.cell(ws.max_row, 2).fill = Y

# ---------------- Теплосеть СПб: НМЦ по позициям нет → цена от себестоимости, итог −45 % к НМЦ (с НДС)
Q = S["spec"]; NMC_S = 83325438.42; TS = r2(NMC_S * 0.55)
ps = fit([(p["cost_unit_vat"], p["qty"], p["cost_unit_vat"] * 1.08) for p in Q], TS, 1.25, 1.65)
rows = []
for p, pr in zip(Q, ps):
    c = p["cost_unit_vat"]
    rows.append([p["n"], p["name"][:70], p["dn"], p["unit"], p["qty"], round(c), round(c * p["qty"]), pr, r2(pr * p["qty"]), round((pr / c - 1) * 100, 1),
                 round(pr * p["qty"] - c * p["qty"]), p.get("cost_class", ""), (p.get("manufacturer_candidate") or "")[:60], (p.get("cost_basis") or "")[:120]])
tc = sum(r[6] for r in rows); tb = sum(r[8] for r in rows)
rows.append(["", "ИТОГО", "", "", sum(p["qty"] for p in Q), "", round(tc), "", r2(tb), round((tb / tc - 1) * 100, 1), round(tb - tc), "", "", f"НМЦ {NMC_S:,.2f} с НДС; заявка −45 %".replace(",", " ")])
ws = sheet("Теплосеть СПб", ["№", "Позиция", "Ду", "Ед.", "Кол-во", "Себестоимость ед. с НДС (оценка)", "Себестоимость сумма", "ЦЕНА ЕД. В ЗАЯВКУ с НДС",
           "Сумма заявки", "Наценка, %", "Валовая прибыль", "Класс оценки (А счёт / Б розница / В оценка)", "Производитель-кандидат", "Основа"], rows,
           [4, 40, 6, 5, 7, 14, 14, 14, 14, 9, 14, 12, 26, 40], {6, 7, 8, 9, 11})
for c in ws[ws.max_row]: c.font = Font(bold=True)
for row in ws.iter_rows(min_row=2, max_row=ws.max_row - 1):
    if row[11].value == "В": row[5].fill = Y
ws.append([]); ws.append(["", "Сценарии (весь объём)", "", "", "", "Цена с НДС", "", "Прибыль (НДС 22 % с вычетом)"])
for lab, k in [("НМЦ", "NMC"), ("−20 %", "-20%"), ("−30 %", "-30%")]:
    sc = S["economics"]["retail_dn_ru"]["scenarios"].get(k)
    if sc: ws.append(["", lab, "", "", "", round(sc["price_vat"]), "", round(sc["profit_vat22"])])
ws.append(["", "НМЦ по позициям заказчик не раскрыл — цены в заявку построены от себестоимости (розница dn.ru/bkarm/СибЗТА, счета ДН.РУ), итог = НМЦ −45 % (цена, при которой обходим опытного конкурента). Жёлтым — себестоимость класса «В» (оценка: электроприводы, Ду600–700 из 09Г2С) — нужны КП заводов."])
ws.cell(ws.max_row, 2).fill = Y
wb.save("../TABLICY_RUSGIDRO_TEPLOSET_2211.xlsx")
print("Нерюнгри: НМЦ", round(NMC_N), "заявка", round(sum(pr*p['qty'] for pr,p in zip(pn,P))), "себест", round(sum(c*p['qty'] for c,p in zip(cu,P))))
print("СПб: НМЦ", NMC_S, "заявка", round(sum(pr*p['qty'] for pr,p in zip(ps,Q))), "себест", round(sum(p['cost_unit_vat']*p['qty'] for p in Q)))
