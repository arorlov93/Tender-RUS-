"""РФ: сводка по всем закупкам с присланными ТЗ (15–16.11.2026) → ../RF_MARZHA_TZ_1611.xlsx + .pdf
Собирает все data/zakupki_1511_tz/*/analysis.json (схема: tender, deadline, nmc, positions, requirements, economics, suppliers, verdict, ...)."""
import json, os, re, glob, html, subprocess
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as CL
from isp_letterhead import SCRATCH
B = "data/zakupki_1511_tz/"
NUM = lambda v: v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def num(v):
    if NUM(v) is not None: return v
    if isinstance(v, dict):
        for k in ("total", "value", "sum", "rub", "amount", "base"):
            if NUM(v.get(k)) is not None: return v[k]
    if isinstance(v, str):
        m = re.search(r"-?\d[\d\s]*[.,]?\d*", v.replace(" ", " "))
        if m:
            try: return float(m[0].replace(" ", "").replace(",", "."))
            except ValueError: return None
    return None


def txt(v, n=400):
    if v is None: return ""
    if isinstance(v, str): return v[:n]
    if isinstance(v, dict): v = v.get("name") or v.get("customer") or v.get("subject") or json.dumps(v, ensure_ascii=False)
    return str(v)[:n]


def number_of(folder, a):
    t = a.get("tender")
    for s in ([t.get(k) for k in ("number", "reg_number", "regNumber", "id")] if isinstance(t, dict) else [t]) + [folder]:
        m = re.search(r"\b(\d{19}|\d{11})\b", str(s or ""))
        if m: return m[1]
    return ""


def verdict_class(v):
    v = v.upper()
    if "NO" in v and "GO" in v: return "NO-GO"
    if "COND" in v or "УСЛОВ" in v: return "CONDITIONAL"
    if "GO" in v or "ПОДАВ" in v: return "GO"
    if "НЕ " in v: return "NO-GO"
    return v[:20] or "—"


VRU = {"GO": "Подавать", "CONDITIONAL": "При условиях", "NO-GO": "Не подавать"}
ROWS = []
for p in sorted(glob.glob(B + "*/analysis.json")):
    folder = p.split("/")[-2]
    try: a = json.load(open(p))
    except Exception as ex: print("skip", p, ex); continue
    e = a.get("economics") or {}
    t = a.get("tender") if isinstance(a.get("tender"), dict) else {}
    n = number_of(folder, a)
    cust = txt(t.get("customer") or t.get("buyer") or a.get("customer"), 80)
    subj = txt(t.get("subject") or t.get("name") or t.get("title") or (a.get("tender") if isinstance(a.get("tender"), str) else ""), 160)
    if not cust and not subj: subj = txt(a.get("tender"), 200)
    v = txt(a.get("verdict"), 200); vc = verdict_class(v)
    nmc = num(a.get("nmc")) or num(e.get("revenue_nmc"))
    ROWS.append(dict(folder=folder, n=n, cust=cust, subj=subj, dl=txt(a.get("deadline") or t.get("deadline"), 40), nmc=nmc,
                     cost=num(e.get("cost")), p0=num(e.get("profit_nmc")), p10=num(e.get("profit_m10")), p20=num(e.get("profit_m20")),
                     m10=num(e.get("margin_pct_m10")), be=num(e.get("breakeven_discount_pct")), vc=vc, reason=txt(a.get("verdict_reason") or v, 700),
                     todo=a.get("todo") or [], risks=a.get("risks") or [], miss=a.get("missing_docs") or [], pos=a.get("positions") or [],
                     sup=a.get("suppliers") or [], a=a))
order = {"GO": 0, "CONDITIONAL": 1, "NO-GO": 2}
ROWS.sort(key=lambda r: (order.get(r["vc"], 3), -(r["p10"] or -1e12)))
link = lambda n: (f"https://zakupki.gov.ru/epz/order/notice/notice223/common-info.html?regNumber={n}" if len(n) == 11 else
                  f"https://zakupki.gov.ru/epz/order/notice/ea20/view/common-info.html?regNumber={n}" if len(n) == 19 else "")
lst = lambda x: x if isinstance(x, list) else [x] if x else []
# ---------- Excel
H = PatternFill("solid", fgColor="0B2545"); HF = Font(bold=True, color="FFFFFF"); W = Alignment(wrap_text=True, vertical="top"); LINK = Font(color="0563C1", underline="single")
wb = Workbook(); ws = wb.active; ws.title = "Сводка"
ws.append(["Номер", "Ссылка ЕИС", "Заказчик", "Предмет", "Подача", "НМЦ, ₽", "Себестоимость, ₽", "Прибыль при НМЦ", "Прибыль −10%", "Маржа −10%, %", "Прибыль −20%", "Безубыточное снижение, %", "Решение", "Почему", "Что сделать", "Риски", "Не хватает документов", "Папка"])
for c in ws[1]: c.fill = H; c.font = HF; c.alignment = W
for r in ROWS:
    ws.append([r["n"], link(r["n"]), r["cust"], r["subj"], r["dl"], r["nmc"], r["cost"], r["p0"], r["p10"], r["m10"], r["p20"], r["be"], VRU.get(r["vc"], r["vc"]), r["reason"],
               "\n".join(map(lambda x: txt(x, 300), lst(r["todo"]))), "\n".join(map(lambda x: txt(x, 300), lst(r["risks"]))), "\n".join(map(lambda x: txt(x, 300), lst(r["miss"]))), r["folder"]])
    if r["n"]: c = ws.cell(ws.max_row, 2); c.hyperlink = c.value; c.font = LINK
    for cc in (6, 7, 8, 9, 11): ws.cell(ws.max_row, cc).number_format = "# ##0"
for i, w in enumerate([20, 30, 28, 44, 16, 14, 14, 13, 13, 9, 13, 10, 14, 60, 60, 50, 40, 24], 1): ws.column_dimensions[CL(i)].width = w
for row in ws.iter_rows(min_row=2):
    for c in row: c.alignment = W
ws.freeze_panes = "C2"; ws.auto_filter.ref = ws.dimensions
keys = ["n", "name", "spec", "equiv_allowed", "unit", "qty", "nmc_unit", "nmc_sum", "supplier_price", "supplier", "url", "note"]
for r in ROWS:
    s = wb.create_sheet((r["n"][-11:] or r["folder"])[:31])
    s.append([f"{r['cust']} — {r['subj']}"]); s.cell(1, 1).font = Font(bold=True)
    s.append(["№", "Наименование", "Требования ТЗ", "Эквивалент", "Ед.", "Кол-во", "НМЦ за ед.", "НМЦ сумма", "Цена поставщика", "Поставщик", "Ссылка на цену", "Примечание"])
    for c in s[2]: c.fill = H; c.font = HF; c.alignment = W
    for p in lst(r["pos"]):
        if not isinstance(p, dict): s.append([txt(p, 800)]); continue
        s.append([json.dumps(p.get(k), ensure_ascii=False)[:800] if isinstance(p.get(k), (dict, list)) else p.get(k) for k in keys])
        u = p.get("url")
        if isinstance(u, str) and u.startswith("http"): c = s.cell(s.max_row, 11); c.hyperlink = u; c.font = LINK
    for title, val in (("Поставщики", r["sup"]), ("Экономика", r["a"].get("economics")), ("Требования", r["a"].get("requirements"))):
        s.append([]); s.append([title]); s.cell(s.max_row, 1).font = Font(bold=True)
        for x in (lst(val) if title == "Поставщики" else [val]): s.append([json.dumps(x, ensure_ascii=False)[:30000]])
    for i, w in enumerate([5, 40, 60, 10, 6, 8, 11, 12, 12, 24, 36, 40], 1): s.column_dimensions[CL(i)].width = w
    for row in s.iter_rows(min_row=2):
        for c in row: c.alignment = W
w3 = wb.create_sheet("Как считано")
for x in ["Прибыль — после УСН 15% (доходы минус расходы, минимум 1% от дохода), без НДС, если в разборе не указано иное (варианты с НДС 5% — на листе закупки, блок «Экономика»).",
          "Себестоимость — цены производителей/дилеров из открытых прайсов и публикаций со ссылками + доставка (оценка). Это не КП: перед подачей запросить КП.",
          "Сортировка: сначала «Подавать», затем «При условиях», затем «Не подавать»; внутри — по прибыли при снижении 10%."]: w3.append([x])
w3.column_dimensions["A"].width = 160
wb.save("../RF_MARZHA_TZ_1611.xlsx")
# ---------- PDF (сводка + коротко по каждой)
e = html.escape; f = lambda v: "—" if v is None else f"{v:,.0f}".replace(",", " ")
css = ("@page{size:A4 landscape;margin:7mm}body{font-family:'DejaVu Sans',Arial;font-size:6.9px;color:#111}h1{font-size:12px;margin:0 0 2px;color:#0B2545}h2{font-size:8.5px;margin:5px 0 1px;color:#0B2545}"
       "table{border-collapse:collapse;width:100%}th{background:#0B2545;color:#fff;padding:2px 3px;text-align:left}td{border-bottom:1px solid #ddd;padding:1px 3px;vertical-align:top}.n{text-align:right;white-space:nowrap}"
       ".g{color:#555}a{color:#0563C1;text-decoration:none}.GO{color:#1B7F3B;font-weight:bold}.CONDITIONAL{color:#B07D00;font-weight:bold}.NO-GO{color:#B00020;font-weight:bold}.neg{color:#B00020}")
h = [f"<!DOCTYPE html><html><head><meta charset=utf-8><style>{css}</style></head><body><h1>Маржа по закупкам РФ с присланными ТЗ — ООО «Гидропас», УСН ({len(ROWS)} закупок, 16.11.2026)</h1>",
     "<div class=g>Прибыль после УСН 15% (мин. 1% от дохода), без НДС. Себестоимость — открытые прайсы + доставка (оценка), не КП. Номер — ссылка на ЕИС. Детали по позициям — в Excel.</div>",
     "<table><tr><th>Номер</th><th>Заказчик / предмет</th><th>Подача</th><th>НМЦ</th><th>Себест.</th><th>При НМЦ</th><th>−10%</th><th>%</th><th>−20%</th><th>Решение</th><th>Главное / чего не хватает</th></tr>"]
for r in ROWS:
    nm = f"<a href='{link(r['n'])}'>{r['n']}</a>" if r["n"] else e(r["folder"])
    miss = "; ".join(txt(x, 90) for x in lst(r["miss"])[:2])
    h.append(f"<tr><td>{nm}</td><td><b>{e(r['cust'])}</b><br>{e(r['subj'][:110])}</td><td>{e(r['dl'][:22])}</td><td class=n>{f(r['nmc'])}</td><td class=n>{f(r['cost'])}</td><td class=n>{f(r['p0'])}</td>"
             f"<td class='n{' neg' if (r['p10'] or 0) < 0 else ''}'><b>{f(r['p10'])}</b></td><td class=n>{'' if r['m10'] is None else round(r['m10'], 1)}</td><td class='n{' neg' if (r['p20'] or 0) < 0 else ''}'>{f(r['p20'])}</td>"
             f"<td class={r['vc']}>{VRU.get(r['vc'], e(r['vc']))}</td><td>{e(r['reason'][:230])}{'<br><span class=g>Нет: ' + e(miss) + '</span>' if miss else ''}</td></tr>")
tot = sum(r["p10"] for r in ROWS if r["vc"] != "NO-GO" and (r["p10"] or 0) > 0)
h.append(f"</table><div class=g>Потенциал (подавать + при условиях, прибыль при −10%): <b>{f(tot)} ₽</b>.</div>")
for r in [x for x in ROWS if x["vc"] != "NO-GO"]:
    h.append(f"<h2>{e(r['n'] or r['folder'])} — {e(r['cust'])}: {VRU.get(r['vc'], r['vc'])}</h2><div><b>Сделать:</b> " + " • ".join(e(txt(x, 160)) for x in lst(r["todo"])[:5]) +
             "<br><b>Риски:</b> " + " • ".join(e(txt(x, 140)) for x in lst(r["risks"])[:3]) + "</div>")
h.append("</body></html>")
hp = f"{SCRATCH}/rf_tz_1611.html"; open(hp, "w").write("".join(h))
subprocess.run(["/opt/pw-browsers/chromium", "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer", f"--print-to-pdf={os.path.abspath('../RF_MARZHA_TZ_1611.pdf')}", "file://" + hp], check=True, capture_output=True, timeout=180)
import pypdf; print("pages", len(pypdf.PdfReader("../RF_MARZHA_TZ_1611.pdf").pages), "rows", len(ROWS), "tot", tot)
for r in ROWS: print(r["vc"], r["n"] or r["folder"], r["nmc"], r["p10"])
