"""РФ: разбор 5 закупок по присланным ТЗ (15.11.2026) → ../RF_RAZBOR_TZ_1511.xlsx + .pdf"""
import json, os, html, subprocess
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as CL
from isp_letterhead import SCRATCH
B = "data/zakupki_1511_tz/"
T = [("kuzbass_32616549223", "32616549223", "Россети Сибирь – Кузбассэнерго-РЭС", "Видеорегистраторы Страж-ПВР-05 64 ГБ, 17 шт (одна позиция, привязка к ПО «Страж-Терминал»)", "18.11 10:00 МСК", "ЗОУ МСП, 223-ФЗ"),
     ("sakh_0361200015026004877", "0361200015026004877", "ГБУЗ ЮСГКБ им. Анкудинова (Сахалин)", "Пробирка вакуумная с активатором свёртывания 4–4,9 мл 13×75, 46 150 шт", "18.11 09:00 Сахалин = 01:00 МСК", "эл. аукцион 44-ФЗ, СМП, запрет ПП 1875"),
     ("hakas_32616563695", "32616563695", "Россети Сибирь – Хакасэнерго", "ОПС/видеонаблюдение: 7 позиций (TRASSIR камеры и MiniNVR, РЕВИЗОР, Страж, БП 12В/18А, АКБ)", "19.11 11:00 МСК", "конкурс МСП, 223-ФЗ"),
     ("kirov_32616573584", "32616573584", "Россети Центр и Приволжье – Кировэнерго", "Расходники/комплектующие СДТУ: 73 позиции (iRZ, Меркурий, Icom, Mean Well, шкафы ЦМО, кабели)", "20.11 11:00 МСК", "запрос предложений МСП, 223-ФЗ"),
     ("tomsk_32616554083", "32616554083", "Россети Томск", "Рамочный каталог 526 позиций связи/энергетики без количеств, лимит 1 874 000 ₽ без НДС", "19.11 11:00 МСК", "запрос предложений, 223-ФЗ")]
ROWS = []
for d, n, cust, what, dl, typ in T:
    a = json.load(open(B + d + "/analysis.json")); e = a["economics"]
    v = str(a.get("verdict")); vr = str(a.get("verdict_reason") or (v if len(v) > 20 else ""))
    ROWS.append(dict(d=d, n=n, cust=cust, what=what, dl=dl, typ=typ, nmc=a.get("nmc"), cost=e.get("cost"), rev=e.get("revenue_nmc"), p0=e.get("profit_nmc"), p10=e.get("profit_m10"), p20=e.get("profit_m20"),
                     m10=e.get("margin_pct_m10"), be=e.get("breakeven_discount_pct"), verdict=v.split(" —")[0][:20], reason=vr, todo=a.get("todo") or [], risks=a.get("risks") or [], pos=a.get("positions") or [], sup=a.get("suppliers") or [], a=a))
link = lambda n: f"https://zakupki.gov.ru/epz/order/notice/notice223/common-info.html?regNumber={n}" if len(n) == 11 else f"https://zakupki.gov.ru/epz/order/extendedsearch/results.html?searchString={n}"
VRU = {"GO": "Подавать", "CONDITIONAL": "Подавать при условиях", "CONDITIONAL GO": "Подавать при условиях", "NO-GO": "Не подавать"}
# ---------- PDF
e = html.escape; f = lambda v: "—" if not isinstance(v, (int, float)) else f"{v:,.0f}".replace(",", " ")
css = ("@page{size:A4 landscape;margin:8mm}body{font-family:'DejaVu Sans',Arial;font-size:7.4px;color:#111}h1{font-size:13px;margin:0 0 3px;color:#0B2545}h2{font-size:9.5px;margin:6px 0 2px;color:#0B2545}"
       "table{border-collapse:collapse;width:100%}th{background:#0B2545;color:#fff;padding:2px 3px;text-align:left}td{border-bottom:1px solid #ddd;padding:2px 3px;vertical-align:top}.n{text-align:right;white-space:nowrap}"
       ".g{color:#555}a{color:#0563C1;text-decoration:none}.go{color:#1B7F3B;font-weight:bold}.c{color:#B07D00;font-weight:bold}.no{color:#B00020;font-weight:bold}ul{margin:0;padding-left:12px}")
h = [f"<!DOCTYPE html><html><head><meta charset=utf-8><style>{css}</style></head><body><h1>Разбор 5 закупок по присланной документации (15.11.2026) — ООО «Гидропас», УСН</h1>",
     "<div class=g>Прибыль — после УСН 15% (мин. 1% от дохода), без НДС (вариант с НДС 5% — в Excel). Себестоимость — цены поставщиков из открытых прайсов со ссылками + доставка (оценка). Цены не являются КП — перед подачей запросить КП.</div>",
     "<table><tr><th>Номер (ссылка)</th><th>Заказчик / что закупают</th><th>Подача</th><th>НМЦ, ₽</th><th>Себестоим., ₽</th><th>Прибыль при НМЦ</th><th>при −10%</th><th>при −20%</th><th>Решение</th></tr>"]
for r in ROWS:
    cls = "go" if r["verdict"] == "GO" else "no" if "NO" in r["verdict"] else "c"
    h.append(f"<tr><td><a href='{link(r['n'])}'>{r['n']}</a><br><span class=g>{e(r['typ'])}</span></td><td><b>{e(r['cust'])}</b><br>{e(r['what'])}</td><td>{e(r['dl'])}</td><td class=n>{f(r['rev'])}</td><td class=n>{f(r['cost'])}</td>"
             f"<td class=n>{f(r['p0'])}</td><td class=n><b>{f(r['p10'])}</b> ({r['m10']}%)</td><td class=n>{f(r['p20'])}</td><td class={cls}>{VRU.get(r['verdict'], r['verdict'])}</td></tr>")
h.append("</table>")
for r in ROWS:
    h.append(f"<h2>{r['n']} — {e(r['cust'])}: {VRU.get(r['verdict'], r['verdict'])}</h2><table><tr><td style='width:50%'><b>Почему:</b> {e(r['reason'][:520])}<br><b>Риски:</b><ul>" + "".join(f"<li>{e(str(x)[:170])}</li>" for x in r["risks"][:4]) + "</ul></td>"
             "<td><b>Что сделать:</b><ul>" + "".join(f"<li>{e(str(x)[:190])}</li>" for x in r["todo"][:6]) + "</ul></td></tr></table>")
h.append("<div class=g>Газпром теплоэнерго МО 32616526616 (23,7 млн ₽, подача 18.11) — документация 5044_KO.rar не получена, не разобрана.</div></body></html>")
hp = f"{SCRATCH}/rf_tz.html"; open(hp, "w").write("".join(h))
subprocess.run(["/opt/pw-browsers/chromium", "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer", f"--print-to-pdf={os.path.abspath('../RF_RAZBOR_TZ_1511.pdf')}", "file://" + hp], check=True, capture_output=True, timeout=180)
import pypdf; print("pages", len(pypdf.PdfReader("../RF_RAZBOR_TZ_1511.pdf").pages))
# ---------- Excel
H = PatternFill("solid", fgColor="0B2545"); HF = Font(bold=True, color="FFFFFF"); W = Alignment(wrap_text=True, vertical="top"); LINK = Font(color="0563C1", underline="single")
wb = Workbook(); ws = wb.active; ws.title = "Сводка"
ws.append(["Номер", "Ссылка", "Заказчик", "Что закупают", "Способ", "Подача", "НМЦ (база оценки), ₽", "Себестоимость, ₽", "Прибыль при НМЦ", "Прибыль −10%", "Маржа −10%, %", "Прибыль −20%", "Безубыточное снижение, %", "Решение", "Почему", "Что сделать", "Риски"])
for c in ws[1]: c.fill = H; c.font = HF; c.alignment = W
for r in ROWS:
    ws.append([r["n"], link(r["n"]), r["cust"], r["what"], r["typ"], r["dl"], r["rev"], r["cost"], r["p0"], r["p10"], r["m10"], r["p20"], r["be"], VRU.get(r["verdict"], r["verdict"]), r["reason"], "\n".join(map(str, r["todo"])), "\n".join(map(str, r["risks"]))])
    c = ws.cell(ws.max_row, 2); c.hyperlink = c.value; c.font = LINK
    for cc in (7, 8, 9, 10, 12): ws.cell(ws.max_row, cc).number_format = "# ##0"
for i, w in enumerate([16, 30, 28, 44, 20, 16, 14, 14, 13, 13, 9, 13, 10, 16, 60, 70, 60], 1): ws.column_dimensions[CL(i)].width = w
for row in ws.iter_rows(min_row=2):
    for c in row: c.alignment = W
for r in ROWS:
    s = wb.create_sheet(r["n"][-11:])
    keys = ["n", "name", "spec", "equiv_allowed", "ktru", "unit", "qty", "nmc_unit", "nmc_sum", "supplier_price", "supplier", "ru_number", "registry", "url", "note"]
    s.append(["№", "Наименование", "Требования ТЗ", "Эквивалент", "КТРУ", "Ед.", "Кол-во", "НМЦ за ед.", "НМЦ сумма", "Цена поставщика", "Поставщик", "РУ", "Реестр", "Ссылка на цену", "Примечание"])
    for c in s[1]: c.fill = H; c.font = HF; c.alignment = W
    for p in r["pos"]:
        s.append([str(p.get(k))[:600] if isinstance(p.get(k), (dict, list)) else p.get(k) for k in keys])
        u = p.get("url")
        if isinstance(u, str) and u.startswith("http"): c = s.cell(s.max_row, 14); c.hyperlink = u; c.font = LINK
    s.append([]); s.append(["Поставщики"]); s.cell(s.max_row, 1).font = Font(bold=True)
    for x in r["sup"]: s.append([json.dumps(x, ensure_ascii=False)[:800]])
    s.append([]); s.append(["Экономика (полностью)"]); s.cell(s.max_row, 1).font = Font(bold=True); s.append([json.dumps(r["a"]["economics"], ensure_ascii=False)[:30000]])
    s.append([]); s.append(["Требования"]); s.cell(s.max_row, 1).font = Font(bold=True); s.append([json.dumps(r["a"]["requirements"], ensure_ascii=False)[:30000]])
    for i, w in enumerate([5, 40, 60, 10, 18, 6, 8, 11, 12, 12, 24, 16, 16, 36, 40], 1): s.column_dimensions[CL(i)].width = w
    for row in s.iter_rows(min_row=2):
        for c in row: c.alignment = W
    s.freeze_panes = "C2"
wb.save("../RF_RAZBOR_TZ_1511.xlsx"); print("ok")
