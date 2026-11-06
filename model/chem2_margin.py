"""Новые химические лоты (поиск ЕИС 06.11.2026, подача с 12.11): маржа при −10 %.
data/companies/chem2_lots.json + chem2_prices_A.json + chem2_prices_B.json → ../HIMIYA_NOVYE_LOTY_0611.xlsx + .pdf"""
import json, os, html, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from pdfutil import html_to_pdf
from tidy import tidy

D = "data/companies/"
LOTS = {x["n"]: x for x in json.load(open(D + "chem2_lots.json"))}
P = {}
for f in ("chem2_prices_A.json", "chem2_prices_B.json"):
    if os.path.exists(D + f):
        for x in json.load(open(D + f)): P[str(x["n"])] = x
CHK = {"32616541193": ("ПРОВЕРИТЬ ТЗ", "НМЦ 44,3 тыс./т между ценой раствора 46 % (26 тыс.) и 100 % NaOH (56 тыс.): если тонны в пересчёте на 100 % — убыток."),
       "32616543663": ("ПРОВЕРИТЬ ТЗ", "НМЦ ≈140 тыс./т без НДС — вдвое выше рынка; вероятно особая марка/фасовка."),
       "32616534976": ("ИДЁМ (проверить марку)", "в ТЗ кристаллический, у источника гранулированный."),
       "32616524946": ("ТОНКО (крупный)", "915 млн, 27 тыс. т — нужен контракт с заводом, кислотные цистерны, банковская гарантия."),
       "32616531354": ("ИДЁМ (нужна квота БСК)", "140 млн, 3 078 т — квота у Башкирской содовой.")}
DIST = {"32616552377": 1900, "32616549952": 1550, "32616549809": 1550, "32616547400": 1550, "32616546151": 800, "32616543474": 1200, "32616526278": 750, "32616543663": 900}
num = lambda v: float(v) if isinstance(v, (int, float)) else None

rows = []
for n, l in LOTS.items():
    p = P.get(n)
    if not p: continue
    nmc = num(l["nmc"]); q = num(p.get("qty")); nu = num(p.get("nmc_unit")) or (nmc / q if nmc and q else None)
    b = num(p.get("buy_unit_vat")); dl = num(p.get("delivery_unit"))
    if dl is None and n in DIST:  # оценка автодоставки: км × 6 ₽/т·км (машина 20 т ≈95 ₽/км + ДОПОГ)
        dl = DIST[n] * 6.0 * (0.001 if str(p.get("unit")) == "кг" else 1); p["delivery_note"] = f"ОЦЕНКА {DIST[n]} км × 6 ₽/т·км. " + str(p.get("delivery_note") or "")
    dl = dl or 0
    vat = str(p.get("nmc_vat") or "неясно")
    m = dict(usn=None, osno=None, p_usn=None)
    if nu and b:
        c_v = b + dl; c_nv = c_v / 1.22
        if vat.startswith("без"):
            bid = nu * .9; usn = (bid - c_v) / bid; osno = (bid - c_nv) / bid
        else:
            bid_v = nu * .9; bid_nv = bid_v / 1.22
            usn = (bid_nv - c_v) / bid_nv if vat.startswith("неясно") else (bid_v - c_v) / bid_v  # «неясно»: худший случай для УСН — сравнение без НДС
            osno = (bid_nv - c_nv) / bid_nv
        m = dict(usn=usn, osno=osno, p_usn=(usn * (nu * .9) * q) if q else None)
    v = m["usn"]
    dec = "нет цены" if v is None else "ИДЁМ" if v >= .15 else "ТОНКО" if v >= .05 else "НЕ ИДЁМ"
    if v is not None and dec != "ИДЁМ" and (m["osno"] or 0) >= .15: dec += " (ОСНО ≥15 %)"
    if n in CHK: dec = CHK[n][0]; p["notes"] = "ПРОВЕРКА: " + CHK[n][1] + " " + str(p.get("notes") or "")
    rows.append(dict(n=n, close=l["close"][:16], cust=l.get("customer"), subj=l["subject"], nmc=nmc, q=q, unit=p.get("unit"), nu=nu, vat=vat, b=b, dl=dl,
                     past=num(p.get("past_unit")), **m, dec=dec, src=p.get("buy_source"), prod=p.get("producers"), notes=p.get("notes"), conf=p.get("confidence"),
                     qsrc=p.get("qty_source"), etp=l.get("etp"), dnote=p.get("delivery_note"), psrc=p.get("past_source")))
ordk = lambda r: (0 if r["dec"].startswith("ИДЁМ") else 1 if r["dec"].startswith(("ТОНКО", "ПРОВЕРИТЬ")) else 2 if r["dec"].startswith("НЕ") else 3, -(r["p_usn"] or 0))
rows.sort(key=ordk)

wb = openpyxl.Workbook(); ws = wb.active; ws.title = f"Химия — {len(rows)} лотов"
H = ["Решение", "№ ЕИС", "Подача до", "Заказчик", "Предмет", "НМЦ, ₽", "Кол-во", "Ед.", "НМЦ за ед.", "НМЦ с/без НДС", "Закупка за ед. с НДС", "Доставка за ед.", "Прошлые контракты за ед.",
     "УСН −10 %", "ОСНО −10 %", "Прибыль УСН −10 %, ₽", "Площадка", "Источник закупочной цены", "Производители", "Доставка", "Риски", "Уверенность", "Откуда кол-во", "Источник прошлых цен"]
ws.append(H)
for r in rows:
    ws.append([r["dec"], r["n"], r["close"], r["cust"], r["subj"], r["nmc"], r["q"], r["unit"], r["nu"], r["vat"], r["b"], r["dl"], r["past"], r["usn"], r["osno"], r["p_usn"],
               r["etp"], r["src"], r["prod"], r["dnote"], r["notes"], r["conf"], r["qsrc"], r["psrc"]])
    X = ws.max_row
    for j in (6, 9, 11, 12, 13, 16): ws.cell(X, j).number_format = "#,##0"
    for j in (14, 15): ws.cell(X, j).number_format = "0%"
    ws.cell(X, 1).fill = PatternFill("solid", fgColor="c6efce" if r["dec"].startswith("ИДЁМ") else "fff2cc" if r["dec"].startswith("ТОНКО") or "ОСНО" in r["dec"] else "f8cbad" if r["dec"].startswith("НЕ") else "f2f2f2")
for c in ws[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
for i, w in enumerate([16, 13, 14, 26, 40, 13, 10, 6, 11, 10, 11, 10, 11, 9, 9, 13, 22, 45, 35, 30, 45, 10, 30, 35], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "C2"; ws.auto_filter.ref = f"A1:X{ws.max_row}"
tidy(wb, row_h=45); wb.save("../HIMIYA_NOVYE_LOTY_0611.xlsx")

e = lambda s: html.escape(str(s if s not in (None, "") else "—"))
fm = lambda v: f"{v:,.0f}".replace(",", " ") if isinstance(v, (int, float)) else "—"
fmm = lambda v: f"{v/1e6:,.2f}".replace(",", " ") if isinstance(v, (int, float)) else "—"
pc = lambda v: f"{v*100:.0f} %" if isinstance(v, (int, float)) else "—"
cls = lambda v: "n" if not isinstance(v, (int, float)) else "ok" if v >= .15 else "warn" if v >= .05 else "bad"
go = [r for r in rows if r["dec"].startswith("ИДЁМ")]
css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:7.3px;color:#111;margin:0}h1{font-size:12px;margin:0 0 3px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #bbb;padding:1px 3px;vertical-align:top}
th{background:#305496;color:#fff;text-align:left}td.n{text-align:right;white-space:nowrap}.ok{background:#c6efce}.warn{background:#fff2cc}.bad{background:#f8cbad}.box{border:1px solid #305496;padding:4px 6px;background:#f3f6fb;margin:3px 0}.s{font-size:6.6px;color:#444}"""
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body><h1>Химия: новые открытые лоты (ЕИС, подача 12–23.11.2026) — маржа при −10 %</h1>",
     f"<div class='box'>Найдено {len(rows)} лотов с оценкой. Маржа ≥15 % на УСН: <b>{len(go)}</b> (прибыль при −10 % ≈{sum(r['p_usn'] or 0 for r in go)/1e6:.1f} млн ₽). "
     "УСН — ГИДРОПАС; где НМЦ «неясно» с НДС или без — считаю худший случай (сравнение без НДС). ОСНО — новая компания. Закупка — цена производителя/дилера с НДС по открытым источникам; КП нет. Количества «у.е./комплект» — из прошлых закупок или «только в документации».</div>",
     "<table><tr><th>Решение</th><th>№ / подача</th><th>Заказчик / предмет</th><th>НМЦ, млн</th><th>Кол-во</th><th>НМЦ за ед.</th><th>Закупка+доставка за ед.</th><th>УСН −10 %</th><th>ОСНО −10 %</th><th>Прибыль УСН, млн</th><th>Риски / источник</th><th>Ув.</th></tr>"]
for r in rows:
    h.append(f"<tr><td class='{cls(r['usn'])}'><b>{e(r['dec'])}</b></td><td class=s>{r['n']}<br>{e(r['close'][:10])}</td><td>{e(str(r['cust'])[:30])}<br><span class=s>{e(r['subj'][:80])}</span></td>"
             f"<td class=n>{fmm(r['nmc'])}</td><td class=n>{fm(r['q'])} {e(r['unit'] or '')}</td><td class=n>{fm(r['nu'])}<br><span class=s>{e(r['vat'])}</span></td><td class=n>{fm(r['b'])} + {fm(r['dl'])}</td>"
             f"<td class='n {cls(r['usn'])}'>{pc(r['usn'])}</td><td class='n {cls(r['osno'])}'>{pc(r['osno'])}</td><td class=n>{fmm(r['p_usn'])}</td><td class=s>{e(str(r['notes'] or '')[:150])}</td><td class=s>{e(r['conf'])}</td></tr>")
h.append("</table></body></html>")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/chem2.html"
open(hp, "w").write("".join(h)); html_to_pdf(hp, "../HIMIYA_NOVYE_LOTY_0611.pdf")
print(len(rows), len(go), [(r["n"], r["dec"], pc(r["usn"]), pc(r["osno"])) for r in rows])
