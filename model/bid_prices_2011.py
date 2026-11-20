"""Позиционные цены для подачи 23.11: Юграэнерго −7,5 % к НМЦ, Энергонефть −9 % к НМЦ, с разбросом по позициям.
→ ../ZAYAVKI_RF_CENY_2011.xlsx"""
import json, random
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

random.seed(2011)
B = "data/zakupki_1511_tz/"
r2 = lambda x: round(x + 1e-9, 2)

def fit(items, target, lo, hi, cap=9e9):
    """items: [(base_unit, qty, floor_unit)] -> unit prices; цена = base*k_i, k_i в [lo,hi], масштаб под target, не ниже floor, не выше cap·base."""
    k = [random.uniform(lo, hi) for _ in items]
    for _ in range(500):
        tot = sum(b * q * ki for (b, q, f), ki in zip(items, k))
        s = target / tot
        k = [min(max(ki * s, f / b if b else 0), cap) for (b, q, f), ki in zip(items, k)]
        if abs(sum(b * q * ki for (b, q, f), ki in zip(items, k)) - target) < 0.5: break
    pr = [r2(b * ki) for (b, q, f), ki in zip(items, k)]
    diff = r2(target - sum(p * q for p, (b, q, f) in zip(pr, items)))
    j = next((i for i, (b, q, f) in enumerate(items) if q == 1), None)
    if j is not None: pr[j] = r2(pr[j] + diff)
    else:  # раскидать копейки по позициям с qty, кратным остатку
        cents = round(diff * 100)
        for i, (b, q, f) in sorted(enumerate(items), key=lambda x: x[1][1]):
            if cents and q == int(q) and cents % int(q) == 0: pr[i] = r2(pr[i] + cents / int(q) / 100); cents = 0
    over = [i for i, (b, q, f) in enumerate(items) if f > b * cap]
    return pr

# ---------- Юграэнерго: НМЦ по позициям нет -> цена от себестоимости (закупка ×1,1 + доставка 60 тыс. пропорц.), насосы Benza 35 300
U = json.load(open(B + "ugra_32616559738/analysis.json"))["positions"]
NMC_U = 1835717.47; TU = r2(NMC_U * 0.925)
goods = sum((35300 if p["n"] == 14 else p["supplier_price"]) * p["qty"] for p in U)
cost_u = [((35300 if p["n"] == 14 else p["supplier_price"] * 1.1)) * (1 + 60000 / goods) for p in U]
pu = fit([(c, p["qty"], c * 1.05) for c, p in zip(cost_u, U)], TU, 1.15, 1.45)

# ---------- Энергонефть: от НМЦ за ед. без НДС, скидка 3–15 % по позициям, итог −9 % (весь объём, гарант + негарант)
E = json.load(open(B + "energoneft_32616571048/analysis.json"))["positions"]
NMC_E = sum(p["nmc_unit_without_vat"] * p["qty_total"] for p in E); TE = r2(NMC_E * 0.91)
FL = {3: 106737.80 / 65}  # счёт ОКПУР
flo = [FL.get(p["n"], p.get("supplier_price_unit") or 0) * 1.03 for p in E]
pe = fit([(p["nmc_unit_without_vat"], p["qty_total"], f) for p, f in zip(E, flo)], TE, 0.85, 0.97, cap=0.99)

wb = Workbook(); H = PatternFill("solid", fgColor="0B2545"); HF = Font(bold=True, color="FFFFFF")
def sheet(title, head, rows, total_row, widths):
    ws = wb.create_sheet(title); ws.append([title]); ws["A1"].font = Font(bold=True, size=12); ws.append(head)
    for c in ws[2]: c.fill = H; c.font = HF; c.alignment = Alignment(wrap_text=True, vertical="top")
    for r in rows: ws.append(r)
    ws.append(total_row)
    for c in ws[ws.max_row]: c.font = Font(bold=True)
    for i, w in enumerate(widths, 1): ws.column_dimensions[chr(64 + i)].width = w
    for row in ws.iter_rows(min_row=3):
        for c in row:
            if isinstance(c.value, float): c.number_format = "# ##0.00"
        row[1].alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "C3"
wb.remove(wb.active)
ru = [[p["n"], p["name"], p["unit"], p["qty"], pr, r2(pr * p["qty"])] for p, c, pr in zip(U, cost_u, pu)]
su = sum(r[5] for r in ru)
sheet("Юграэнерго 32616559738", ["№", "Наименование", "Ед.", "Кол-во", "ЦЕНА ЕД. В ЗАЯВКУ, ₽ (НДС не облагается)", "Сумма, ₽"], ru,
      ["", f"ИТОГО (НМЦ {NMC_U:,.2f} −7,5 %)".replace(",", " "), "", "", "", r2(su)], [5, 50, 10, 8, 18, 16])
re_ = [[p["n"], p["name"], p["unit"], p["qty_total"], p["nmc_unit_without_vat"], pr, round((1 - pr / p["nmc_unit_without_vat"]) * 100, 1), r2(pr * p["qty_total"]), r2(pr * p["qty_guaranteed"])] for p, pr in zip(E, pe)]
se = sum(r[7] for r in re_)
sheet("Энергонефть 32616571048", ["№", "Наименование", "Ед.", "Кол-во всего", "НМЦ ед. без НДС, ₽", "ЦЕНА ЕД. В ЗАЯВКУ, ₽ (НДС не облагается)", "Скидка, %", "Сумма весь объём, ₽", "Сумма гарант. объём, ₽"], re_,
      ["", f"ИТОГО (НМЦ без НДС {NMC_E:,.2f} −9 %)".replace(",", " "), "", "", r2(NMC_E), "", round((1 - se / NMC_E) * 100, 2), r2(se), r2(sum(r[8] for r in re_))], [5, 50, 8, 9, 13, 16, 9, 15, 15])
wb.save("../ZAYAVKI_RF_CENY_2011.xlsx")
print("Югра", r2(su), "цель", TU, "скидка", round((1 - su / NMC_U) * 100, 3), )
print("Энергонефть", r2(se), "цель", TE, "скидка", round((1 - se / NMC_E) * 100, 3), "скидки мин/макс", min(r[6] for r in re_), max(r[6] for r in re_), "гарант", r2(sum(r[8] for r in re_)))
