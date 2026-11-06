"""Итог: реальные цены и маржа по 13 приоритетным лотам (06.11.2026). data/companies/prices_A.json + prices_B.json → ../ITOG_CENY_MARZHA_0611.xlsx + .pdf"""
import json, os, html, ast, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from pdfutil import html_to_pdf
from tidy import tidy

D = "data/companies/"
LOTS = []
for f in ("prices_A.json", "prices_B.json"):
    if os.path.exists(D + f): LOTS += json.load(open(D + f))["lots"]
base = {r[1]: r for r in openpyxl.load_workbook("../KONKURSY_KP_0611.xlsx")["Конкурсы (13)"].iter_rows(min_row=2, values_only=True)}
def P(v):
    if isinstance(v, str):
        try: return ast.literal_eval(v)
        except Exception: return v
    return v
def num(v):
    try: return float(v)
    except Exception: return None
def T(v):
    v = P(v)
    if isinstance(v, dict): return "; ".join(f"{k}: {T(x)}" for k, x in v.items())
    if isinstance(v, list): return "; ".join(T(x) for x in v)
    return "" if v in (None, "None") else str(v)
rows = []
got = set()
for l in LOTS:
    n = str(l["number"]); got.add(n); b = base.get(n, [None] * 19); m = P(l.get("margin")) or {}
    rows.append(dict(n=n, close=str(b[0])[:16] if b[0] else "", cust=b[2], product=T(l.get("product"))[:140], qty=T(l.get("qty")) if T(l.get("qty")) else T(l.get("qty_estimate"))[:40],
                     unit=T(l.get("unit")), nmc=num(l.get("nmc")), nmc_unit=num(l.get("nmc_unit")), buy=num(l.get("best_buy_unit")), buy_note=T(l.get("best_buy_note"))[:300],
                     cost=num(l.get("cost_vat")), deliv=num(l.get("delivery_rub")), m=m, pp=T(l.get("pp719"))[:250], risks=T(l.get("risks"))[:500], verdict=T(l.get("verdict")),
                     conf=T(l.get("confidence")), vat=T(l.get("nmc_vat"))[:160], market=P(l.get("market")) or [], past=P(l.get("past_contracts")) or [], etp=b[10]))
for n, b in base.items():
    if n not in got:
        rows.append(dict(n=n, close=str(b[0])[:16], cust=b[2], product=b[3], qty=b[4], unit=b[5], nmc=b[6], nmc_unit=b[7], buy=None, buy_note="", cost=None, deliv=None, m={}, pp="", risks="",
                         verdict="цены проверяются", conf="", vat="", market=[], past=[], etp=b[10]))
g = lambda r, k: r["m"].get(k) if isinstance(r["m"], dict) else None
def rank(r):
    v = g(r, "usn_m10")
    return -(v if isinstance(v, (int, float)) else -9)
rows.sort(key=rank)
def cls(v):
    if not isinstance(v, (int, float)): return "n"
    return "ok" if v >= 0.15 else "warn" if v > 0 else "bad"

wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Итог (13 лотов)"
H = ["№ ЕИС", "Подача до", "Заказчик", "Товар", "Кол-во", "Ед.", "НМЦ, ₽", "НМЦ за ед.", "Лучшая закупка за ед. (с НДС)", "Доставка, ₽", "Себестоимость с НДС, ₽",
     "УСН (ГИДРОПАС): маржа по НМЦ", "УСН −10 %", "УСН −20 %", "ОСНО: маржа по НМЦ", "ОСНО −10 %", "ОСНО −20 %", "Как сравнивает заказчик", "Вывод", "Уверенность", "Реестр ПП 719", "Риски", "Комментарий к цене закупки", "Площадка"]
ws.append(H)
for r in rows:
    ws.append([r["n"], r["close"], r["cust"], r["product"], r["qty"], r["unit"], r["nmc"], r["nmc_unit"], r["buy"], r["deliv"], r["cost"],
               g(r, "usn_nmc"), g(r, "usn_m10"), g(r, "usn_m20"), g(r, "osno_nmc"), g(r, "osno_m10"), g(r, "osno_m20"), r["vat"], r["verdict"], r["conf"], r["pp"], r["risks"], r["buy_note"], r["etp"]])
    X = ws.max_row
    for j in (7, 8, 9, 10, 11): ws.cell(X, j).number_format = "#,##0"
    for j in range(12, 18):
        ws.cell(X, j).number_format = "0%"
        v = ws.cell(X, j).value
        if isinstance(v, (int, float)): ws.cell(X, j).fill = PatternFill("solid", fgColor={"ok": "c6efce", "warn": "fff2cc", "bad": "f8cbad"}[cls(v)])
for c in ws[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
for i, w in enumerate([13, 15, 22, 40, 12, 8, 13, 12, 13, 11, 13, 10, 9, 9, 10, 9, 9, 30, 45, 10, 40, 60, 50, 24], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "E2"; ws.auto_filter.ref = f"A1:X{ws.max_row}"
s2 = wb.create_sheet("Цены продавцов"); s2.append(["№ ЕИС", "Товар", "Продавец", "Цена за ед. с НДС", "НДС", "Дата", "Ссылка"])
for r in rows:
    for x in r["market"]:
        if isinstance(x, dict): s2.append([r["n"], r["product"][:60], x.get("seller"), num(x.get("price_unit_vat")), x.get("vat_note"), x.get("date"), x.get("url")])
s3 = wb.create_sheet("Прошлые контракты"); s3.append(["№ ЕИС", "Товар", "Заказчик", "Дата", "Цена за ед.", "Источник"])
for r in rows:
    for x in r["past"]:
        if isinstance(x, dict): s3.append([r["n"], r["product"][:60], x.get("customer"), x.get("date"), num(x.get("price_unit")), T(x.get("source"))[:300]])
for s in (s2, s3):
    for c in s[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")
    for i in range(1, 8): s.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 30
    for row in s.iter_rows(min_row=2):
        if isinstance(row[3].value, (int, float)) or isinstance(row[4].value, (int, float)): pass
tidy(wb, row_h=40); wb.save("../ITOG_CENY_MARZHA_0611.xlsx")

e = lambda s: html.escape(str(s if s is not None else "—"))
fm = lambda v: f"{v:,.0f}".replace(",", " ") if isinstance(v, (int, float)) else "—"
pc = lambda v: f"{v*100:.0f} %" if isinstance(v, (int, float)) else "—"
css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:7.8px;color:#111;margin:0}h1{font-size:13px;margin:0 0 3px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #bbb;padding:1px 3px;vertical-align:top}
th{background:#305496;color:#fff;text-align:left}td.n{text-align:right;white-space:nowrap}.ok{background:#c6efce}.warn{background:#fff2cc}.bad{background:#f8cbad}.box{border:1px solid #305496;padding:4px 6px;background:#f3f6fb;margin:4px 0}.s{font-size:7px;color:#444}"""
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body><h1>Реальные цены и маржа по 13 приоритетным лотам (06.11.2026)</h1>",
     "<div class='box'>Цены — прайсы производителей/дилеров и прошлые контракты (источники и даты — в Excel). Маржа — (цена − себестоимость с доставкой) / цена; <b>УСН</b> — для ГИДРОПАС сейчас (если заказчик сравнивает без НДС — наша цена без НДС, закупка с НДС); <b>ОСНО</b> — для новой компании. "
     "Т Плюс (Киров, АН-31, БАУ-А) сравнивает <b>с НДС</b> — там УСН не теряет; Т Плюс Саратов — без НДС; ТГК-1/Форвард — не подтверждено (принято без НДС). Документацию лотов скачать не удалось — часть количеств/марок оценочные (см. «Уверенность»).</div>",
     "<table><tr><th>№</th><th>Подача</th><th>Заказчик</th><th>Товар</th><th>НМЦ</th><th>Закупка за ед.</th><th>Себест.</th><th>УСН НМЦ</th><th>УСН −10%</th><th>УСН −20%</th><th>ОСНО −10%</th><th>Вывод</th><th>Ув.</th></tr>"]
for r in rows:
    h.append(f"<tr><td class=s>{r['n']}</td><td>{e(r['close'][:10])}</td><td>{e(str(r['cust'])[:22])}</td><td>{e(r['product'][:70])}</td><td class=n>{fm(r['nmc'])}</td><td class=n>{fm(r['buy'])}</td><td class=n>{fm(r['cost'])}</td>"
             + "".join(f"<td class='n {cls(g(r, k))}'>{pc(g(r, k))}</td>" for k in ("usn_nmc", "usn_m10", "usn_m20", "osno_m10")) + f"<td class=s>{e(r['verdict'][:150])}</td><td class=s>{e(r['conf'])}</td></tr>")
h.append("</table></body></html>")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/itog.html"
open(hp, "w").write("".join(h)); html_to_pdf(hp, "../ITOG_CENY_MARZHA_0611.pdf")
print(len(rows), len(got))
