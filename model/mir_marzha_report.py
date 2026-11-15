"""PDF + Excel по детальной марже мировых тендеров (из data/tenders_1511/mir_marzha.json)."""
import json, os, html, subprocess
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as CL
from isp_letterhead import SCRATCH
T = "data/tenders_1511/"
D = json.load(open(T + "mir_marzha.json")); A = D["assumptions"]; ROWS = D["rows"]
CB = json.load(open(T + "cost_benchmarks.json"))
ORD = {"Напрямую (иностранцы допускаются)": 0, "Уточнить": 1, "Через местного партнёра / регистрацию": 2}
def ak(o): return ORD.get(o["access"], 3)
SHORT = {"Напрямую (иностранцы допускаются)": "напрямую", "Уточнить": "уточнить", "Через местного партнёра / регистрацию": "через партнёра"}
def sh(o): return SHORT.get(o["access"], "закрыт")
calc = sorted([o for o in ROWS if o["margin_usd"] is not None], key=lambda o: (ak(o), -(o["margin_usd"] or 0)))
be = sorted([o for o in ROWS if o["margin_usd"] is None and o["cost_usd"]], key=lambda o: (ak(o), -(o["qty_t"] or 0)))
rest = [o for o in ROWS if o["margin_usd"] is None and not o["cost_usd"]]
def items_txt(o, n=3):
    return "; ".join(f"{str(l.get('item'))[:38]} {l['qty_t']:g} т" for l in o["lines"][:n] if l.get("qty_t")) or ""

# ---------------- PDF
e = html.escape; k = lambda v: "—" if v is None else f"{v/1e3:,.0f}".replace(",", " ")
grp = {}
for o in calc:
    g = grp.setdefault(sh(o), [0, 0, 0, 0]); g[0] += 1; g[1] += o["rev_usd"] or 0; g[2] += o["cost_usd"] or 0; g[3] += o["margin_usd"]
css = ("@page{size:A4 landscape;margin:7mm}body{font-family:'DejaVu Sans',Arial;font-size:6.5px;color:#111}h1{font-size:12.5px;margin:0 0 2px;color:#0B2545}h2{font-size:9px;margin:5px 0 1px;color:#0B2545}"
       "table{border-collapse:collapse;width:100%}th{background:#0B2545;color:#fff;padding:1px 3px;text-align:left}td{border-bottom:1px solid #ddd;padding:1px 3px;vertical-align:top;line-height:1.15}"
       ".n{text-align:right;white-space:nowrap}.g{color:#555}a{color:#0563C1;text-decoration:none}.w{color:#9a6700}.neg{color:#B00020;font-weight:bold}.pos{color:#1B7F3B;font-weight:bold}")
h = [f"<!DOCTYPE html><html><head><meta charset=utf-8><style>{css}</style></head><body><h1>Зарубежные тендеры ISP Group: спецификации и маржа (расчёт 15.11.2026)</h1>",
     f"<div class=g>Разобрано {len(ROWS)} тендеров (документация получена по {sum(1 for o in ROWS if o['docs'])}). Маржа посчитана по {len(calc)}: выручка — опубликованная стоимость или цена прошлых контрактов того же заказчика на тот же продукт; "
     f"себестоимость — FOB завода по марке из ТЗ + морской фрахт до порта (индексы Drewry/SCFI/FBX, ноябрь 2026) + пошлина страны (TARIC/UKGT/WITS, антидемпинг ЕС на PBTC) + доставка до объекта ${A['inland_eu']:.0f}/т (ЕС, порт) / ${A['inland_far']:.0f}/т (вглубь) + {A['other_pct']*100:.0f}% прочее (допущения). Суммы — тыс. USD.</div>",
     "<table style='width:62%;margin:3px 0'><tr><th>Доступ</th><th>Тендеров</th><th>Выручка</th><th>Себестоимость</th><th>Маржа</th></tr>"]
for g in ["напрямую", "уточнить", "через партнёра", "закрыт"]:
    if g in grp: v = grp[g]; h.append(f"<tr><td>{g}</td><td class=n>{v[0]}</td><td class=n>{k(v[1])}</td><td class=n>{k(v[2])}</td><td class=n><b>{k(v[3])}</b></td></tr>")
h.append("</table><h2>1. Маржа посчитана</h2><table><tr><th>#</th><th>Доступ</th><th>Страна / номер</th><th>Что закупают (по ТЗ)</th><th>Объём, т</th><th>Выручка</th><th>Себест.</th><th>Маржа</th><th>%</th><th>Откуда выручка / риск</th><th>Подача</th></tr>")
for i, o in enumerate(calc, 1):
    num = f"<a href='{e(o['link'])}'>{e(str(o['n'])[:20])}</a>" if o.get("link") else e(str(o["n"])[:20])
    cls = "neg" if o["margin_usd"] < 0 else "pos"
    h.append(f"<tr><td>{i}</td><td>{sh(o)}</td><td>{e(o['country'] or '')}<br>{num}</td><td>{e(o['ru'][:95])}<br><span class=g>{e(items_txt(o))}</span></td><td class=n>{o['qty_t']:,.0f}</td>"
             f"<td class=n>{k(o['rev_usd'])}</td><td class=n>{k(o['cost_usd'])}</td><td class='n {cls}'>{k(o['margin_usd'])}</td><td class=n>{o['margin_pct']:.0f}</td>"
             f"<td class=w>{e(o['rev_src'][:55])}{(' ' + e(o['warn'][:70])) if o.get('warn') else ''}</td><td style='white-space:nowrap'>{e(str(o['dl'] or '')[:11])}</td></tr>")
h.append("</table><h2>2. Себестоимость посчитана, цены заказчика нет — наша минимальная цена (безубыточность, USD/т с доставкой)</h2><table><tr><th>#</th><th>Доступ</th><th>Страна / номер</th><th>Что закупают</th><th>Объём, т</th><th>Себест., тыс.</th><th>Безубыт., $/т</th><th>Подача</th></tr>")
for i, o in enumerate(be, 1):
    num = f"<a href='{e(o['link'])}'>{e(str(o['n'])[:20])}</a>" if o.get("link") else e(str(o["n"])[:20])
    h.append(f"<tr><td>{i}</td><td>{sh(o)}</td><td>{e(o['country'] or '')} {num}</td><td>{e(o['ru'][:110])}</td><td class=n>{(o['qty_t'] or 0):,.0f}</td><td class=n>{k(o['cost_usd'])}</td><td class=n><b>{o['breakeven_t'] or '—'}</b></td><td>{e(str(o['dl'] or '')[:11])}</td></tr>")
h.append(f"</table><div class=g style='margin-top:3px'>Ещё {len(rest)} тендеров без объёма и цены в открытом доступе (документы за логином/платные/выдаются лично) — в Excel, лист «Маржа детально», с причиной. "
         "Ключевые оговорки: ЕС — антидемпинг на PBTC из КНР 192%; США — +12,5% Section 301 на товары КНР; жидкие PAC/сульфат алюминия считаются через пересчёт по Al2O3 — возить раствор из Китая невыгодно, нужен местный партнёр. Цены FOB — индексы/листинги (не КП), кроме PAC, фосфонатов, смол (КП заводов).</div></body></html>")
hp = f"{SCRATCH}/mir_marzha2.html"; open(hp, "w").write("".join(h))
subprocess.run(["/opt/pw-browsers/chromium", "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer", f"--print-to-pdf={os.path.abspath('../ZAKUPKI_MIR_MARZHA_1511.pdf')}", "file://" + hp], check=True, capture_output=True, timeout=180)
import pypdf; print("PDF pages", len(pypdf.PdfReader("../ZAKUPKI_MIR_MARZHA_1511.pdf").pages), "calc", len(calc), "be", len(be), "rest", len(rest), {g: round(v[3]) for g, v in grp.items()})

# ---------------- Excel
wb = load_workbook("../ZAKUPKI_MIR_1511.xlsx")
for n in list(wb.sheetnames):
    if n.startswith(("4.", "5.", "6.", "7.")): del wb[n]
H = PatternFill("solid", fgColor="0B2545"); HF = Font(bold=True, color="FFFFFF"); W = Alignment(wrap_text=True, vertical="top"); LINK = Font(color="0563C1", underline="single")
ws = wb.create_sheet("4. Маржа детально")
cols = ["№", "Доступ", "Страна", "Номер", "Заказчик", "Что закупают (по ТЗ)", "Позиции из ТЗ (продукт — объём — себестоимость с доставкой $/т — цена контрактов $/т)", "Объём, т", "Выручка, USD", "Откуда выручка",
        "Себестоимость, USD", "Маржа, USD", "Маржа, %", "Безубыточность, $/т", "Предупреждение", "Срок договора", "Поставка", "Сертификаты", "Обеспечение", "Реально ли / риски", "Подача", "Ссылка", "Документы"]
ws.append(cols)
for c in ws[1]: c.fill = H; c.font = HF; c.alignment = W
allrows = calc + be + rest
for i, o in enumerate(allrows, 1):
    it = "\n".join(f"• {str(l.get('item'))[:70]} — {l['qty_t']:g} т — " + (f"{l['landed_t']} $/т (FOB {l['fob']} + фрахт {l['freight']} + пошлина {l['duty_pct']}% + доставка {l['inland']:.0f})" if l.get("landed_t") else l.get("note", "")) + (f" — контракты {l['bench_t']} $/т" if l.get("bench_t") else "") for l in o["lines"])
    ws.append([i, o["access"], o["country"], o["n"], o["buyer"], o["ru"], it, o["qty_t"], o["rev_usd"], o["rev_src"], o["cost_usd"], o["margin_usd"], (o["margin_pct"] or 0) / 100 if o["margin_pct"] is not None else None, o["breakeven_t"],
               o.get("warn"), str(o.get("period") or "")[:300], str(o.get("delivery") or "")[:300], str(o.get("certs") or "")[:300], str(o.get("security") or "")[:200], (str(o.get("feas") or "") + " | " + str(o.get("risks") or ""))[:600], o["dl"], None, "\n".join(o.get("doc_urls") or [])[:500]])
    if o.get("link"): c = ws.cell(ws.max_row, 22); c.value = o["link"]; c.hyperlink = o["link"]; c.font = LINK
    for cc in (9, 11, 12, 14): ws.cell(ws.max_row, cc).number_format = "# ##0"
    ws.cell(ws.max_row, 13).number_format = "0%"
for row in ws.iter_rows(min_row=2):
    for c in row: c.alignment = W
for i, w in enumerate([4, 16, 12, 18, 26, 44, 70, 9, 13, 30, 13, 13, 8, 11, 30, 22, 26, 22, 18, 50, 12, 36, 36], 1): ws.column_dimensions[CL(i)].width = w
ws.freeze_panes = "E2"; ws.auto_filter.ref = ws.dimensions
w5 = wb.create_sheet("5. Допущения и источники")
w5.append(["Что", "Значение", "Источник / комментарий"])
for x in [("Загрузка 20' контейнера", f"{A['payload_t']} т", "допущение (IBC/мешки 20–22 т)"), ("Доставка от порта до объекта, ЕС/портовые страны", f"${A['inland_eu']}/т", "ДОПУЩЕНИЕ; ориентир: автодоставка ЕС €2,7–3,0/км (Upply/Ti/IRU Q3-2026)"),
          ("Доставка вглубь континента", f"${A['inland_far']}/т", "ДОПУЩЕНИЕ"), ("Прочие расходы (банк, страховка, инспекция)", f"{A['other_pct']*100:.0f}%", "ДОПУЩЕНИЕ"), ("Надбавка DG класс 8", f"${A['dg']}/конт.", "Hapag-Lloyd DGP")]: w5.append(list(x))
w5.append([]); w5.append(["Цены FOB (USD/т)", "", ""])
for p in CB["products"]: w5.append([p.get("name_ru") or p["key"], p.get("fob_typical"), f"{p.get('source_type') or ''} {p.get('source_url')} ({p.get('source_date')})"])
w5.append([]); w5.append(["Фрахт", "", ""])
for f in CB["freight"]: w5.append([f"{f['from']} → {f['to_port']} {f['container']}", f.get("usd_typical"), f"{f.get('source_url')} ({f.get('source_date')})"])
for c in w5[1]: c.fill = H; c.font = HF
w5.column_dimensions["A"].width = 60; w5.column_dimensions["B"].width = 14; w5.column_dimensions["C"].width = 110
wb.save("../ZAKUPKI_MIR_1511.xlsx"); print("xlsx ok", len(allrows))
