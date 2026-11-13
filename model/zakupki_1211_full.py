"""Полный рейтинг закупок 13.11.2026 (подача ≥16.11): пакеты A (пробирки), B (медизделия), C (видео/ИТ/прочее), D (без НМЦ) + ранее разобранные.
→ ../ZAKUPKI_1311_POLNYI_RAZBOR.xlsx + .pdf"""
import json, html, os, subprocess, re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from isp_letterhead import SCRATCH
D = "data/zakupki_1211/"
def L(f):
    d = json.load(open(D + f)); return d if isinstance(d, list) else (d.get("lots") or next(v for v in d.values() if isinstance(v, list)))
def g(o, *p):
    for k in p:
        o = o.get(k) if isinstance(o, dict) else None
    return o
def dl(x):
    s = str(x.get("deadline_msk") or x.get("deadline_eis") or x.get("deadline") or "")
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})", s)
    if m: return f"{m[3]}.{m[2]} {m[4]}:{m[5]}"
    m = re.search(r"(\d{2})\.(\d{2})(?:\.\d{4})? ?(\d{2}:\d{2})?", s)
    return f"{m[1]}.{m[2]} {m[3] or ''}".strip() if m else s[:16]
rows = {}
def add(x, pack, mkey="margin_base"):
    m = x.get(mkey) or x.get("margin") or {}
    go = str(x.get("go") or x.get("verdict") or x.get("cost_status") or "")
    rec = dict(num=str(x["num"]), pack=pack, name=str(x.get("name"))[:120], region=str(x.get("region") or "")[3:], cust=str(x.get("customer") or "")[:70],
               law=x.get("law"), method=x.get("method") or x.get("method_list"), price_only="цена" if (x.get("price_only", True) and "конкурс" not in str(x.get("method")).lower()) else "конкурс",
               nmc=x.get("nmc"), dl=dl(x), p0=g(m, "nmc", "net_profit"), p10=g(m, "m10", "net_profit"), p20=g(m, "m20", "net_profit"),
               pct10=g(m, "m10", "net_margin_pct"), go=go[:90], why=str(x.get("go_reason") or x.get("calc_note") or "; ".join(x.get("risks", [])[:2]) if isinstance(x.get("risks"), list) else x.get("go_reason") or "")[:200],
               cons10=g(x, "margin_conservative", "m10", "net_profit"))
    if rec["num"] not in rows or pack in ("A", "B", "C"): rows[rec["num"]] = rec
for x in L("analysis_xlsx_shortlist.json"): add(x, "S")
for x in L("full_C_result.json"): add(x, "C")
for x in L("full_B_med_result.json"): add(x, "B")
for x in L("full_A_tubes_result.json"): add(x, "A")
TERM = json.load(open(D + "terms.json"))
for r in rows.values(): r["term"] = TERM.get(r["num"])
R = sorted(rows.values(), key=lambda r: -(r["p10"] if isinstance(r["p10"], (int, float)) else -9e12))
def cls(r):
    v = r["p10"]; go = r["go"].upper()
    if "NO-GO" in go or "НЕТ" == go[:3]: return "Нет"
    if not isinstance(v, (int, float)): return "Нет данных"
    if v >= 30000 and (r["pct10"] or 0) >= 8: return "Условно" if ("УСЛ" in go or "условно" in r["go"]) else "GO"
    return "Нет"
for r in R: r["cls"] = cls(r)
DD = L("full_D_result.json")
cnt = {k: sum(1 for r in R if r["cls"] == k) for k in ("GO", "Условно", "Нет", "Нет данных")}
tot_go = sum(r["p10"] for r in R if r["cls"] == "GO"); tot_c = sum(r["p10"] for r in R if r["cls"] == "Условно")

wb = Workbook(); H = PatternFill("solid", fgColor="0B2545"); HF = Font(bold=True, color="FFFFFF"); W = Alignment(wrap_text=True, vertical="top")
ws = wb.active; ws.title = "С НМЦ (рейтинг)"
cols = ["№", "Номер", "Пакет", "Предмет", "Регион", "Заказчик", "Закон", "Способ", "Отбор", "НМЦ, ₽", "Подача (МСК)", "Прибыль при НМЦ, ₽", "при −10%, ₽", "при −20%, ₽", "Маржа −10%, %", "−10% по рознице МиниМед, ₽", "Решение", "Вердикт/риск из анализа", "Срок исполнения, мес"]
ws.append(cols)
for i, r in enumerate(R, 1):
    ws.append([i, r["num"], r["pack"], r["name"], r["region"], r["cust"], r["law"], r["method"], r["price_only"], r["nmc"], r["dl"], r["p0"], r["p10"], r["p20"], r["pct10"], r["cons10"], r["cls"], (r["go"] + " | " + r["why"])[:300], r["term"]])
ws2 = wb.create_sheet("Без НМЦ (D)")
ws2.append(["Номер", "Предмет", "Заказчик", "Регион", "Площадка", "Подача", "Оценка суммы, ₽", "Класс", "Причина", "Что нужно для КП", "Позиции"])
for x in sorted(DD, key=lambda x: ({"стоит": 0, "возможно": 1}.get(x.get("class"), 2), -(x.get("est_size_rub") or 0))):
    ws2.append([x["num"], x.get("name"), x.get("customer"), str(x.get("region") or "")[3:], x.get("platform") or x.get("etp"), dl(x), x.get("est_size_rub"), x.get("class"), x.get("reason"), x.get("to_quote"), str(x.get("items") or x.get("positions"))[:400]])
for s in (ws, ws2):
    for c in s[1]: c.fill = H; c.font = HF; c.alignment = W
    for row in s.iter_rows(min_row=2):
        for c in row: c.alignment = W
for i, w in enumerate([5, 22, 6, 45, 18, 30, 9, 16, 8, 13, 12, 13, 13, 13, 9, 13, 10, 60]): ws.column_dimensions[chr(65 + i)].width = w
for i, w in enumerate([16, 40, 30, 18, 14, 12, 14, 10, 40, 50, 60]): ws2.column_dimensions[chr(65 + i)].width = w
ws.freeze_panes = "E2"; ws2.freeze_panes = "C2"
ws3 = wb.create_sheet("Итог")
for a in [["Лотов с НМЦ разобрано", len(R)], ["GO", cnt["GO"]], ["Условно", cnt["Условно"]], ["Нет", cnt["Нет"]], ["Нет данных", cnt["Нет данных"]],
          ["Прибыль GO при −10%, ₽", tot_go], ["Прибыль условных при −10%, ₽", tot_c], ["Лотов без НМЦ", len(DD)],
          ["Без НМЦ: стоит", sum(1 for x in DD if x.get("class") == "стоит")], ["Без НМЦ: возможно", sum(1 for x in DD if x.get("class") == "возможно")],
          ["Допущения", "УСН 15% (мин. 1%), медизделия без НДС; прочее — НДС 5% без вычета (если доход 2025 >20 млн); логистика, гарантии, оборотка учтены; себестоимость пробирок — BERIMED/МиниМед"]]: ws3.append(a)
wb.save("../ZAKUPKI_1311_POLNYI_RAZBOR.xlsx")

e = html.escape; f = lambda v: "—" if not isinstance(v, (int, float)) else f"{v/1e3:,.0f}".replace(",", " ")
top = [r for r in R if r["cls"] in ("GO", "Условно")]
css = ("@page{size:A4 landscape;margin:8mm}body{font-family:'DejaVu Sans',Arial;font-size:7.6px;color:#111}h1{font-size:13px;margin:0;color:#0B2545}h2{font-size:10px;margin:6px 0 2px;color:#0B2545}"
       "table{border-collapse:collapse;width:100%}th{background:#0B2545;color:#fff;padding:2px 3px;text-align:left}td{border-bottom:1px solid #ddd;padding:2px 3px;vertical-align:top}.n{text-align:right;white-space:nowrap}.g{color:#555}.go{color:#1B7F3B;font-weight:bold}.c{color:#B07D00;font-weight:bold}")
h = [f"<!DOCTYPE html><html><head><meta charset=utf-8><style>{css}</style></head><body><h1>Полный разбор закупок с подачей ≥16.11.2026</h1>",
     f"<div class=g>Лотов с НМЦ: {len(R)} — GO {cnt['GO']}, условно {cnt['Условно']}, нет {cnt['Нет']}, нет данных {cnt['Нет данных']}. Прибыль при −10% (после налога): GO ≈ <b>{tot_go/1e6:.1f} млн ₽</b>, условные ≈ {tot_c/1e6:.1f} млн ₽. "
     f"Без НМЦ: {len(DD)} (стоит {sum(1 for x in DD if x.get('class')=='стоит')}, возможно {sum(1 for x in DD if x.get('class')=='возможно')}). Почти все лоты — отбор только по цене. Себестоимость пробирок — прайс BERIMED/МиниМед; нужны письма производителей о соответствии ТЗ и реестр ПП 719.</div>",
     "<h2>GO и условные (по прибыли при −10%)</h2><table><tr>" + "".join(f"<th>{c}</th>" for c in ["#", "Номер", "Предмет", "Регион", "НМЦ, тыс.", "Подача", "Приб. НМЦ", "−10%", "−20%", "% −10", "−10% МиниМед", "Решение"]) + "</tr>"]
for i, r in enumerate(top, 1):
    h.append(f"<tr><td>{i}</td><td>{e(r['num'])}</td><td>{e(r['name'][:70])}</td><td>{e(r['region'][:18])}</td><td class=n>{f(r['nmc'])}</td><td>{e(r['dl'])}</td>"
             f"<td class=n>{f(r['p0'])}</td><td class=n><b>{f(r['p10'])}</b></td><td class=n>{f(r['p20'])}</td><td class=n>{r['pct10'] if r['pct10'] is not None else '—'}</td><td class=n>{f(r['cons10'])}</td><td class={'go' if r['cls']=='GO' else 'c'}>{r['cls']}</td></tr>")
h.append("</table><h2>Без НМЦ — стоит подавать</h2><table><tr><th>Номер</th><th>Предмет</th><th>Заказчик</th><th>Подача</th><th>Оценка, млн</th><th>Что нужно</th></tr>")
for x in [x for x in DD if x.get("class") == "стоит"]:
    h.append(f"<tr><td>{e(x['num'])}</td><td>{e(str(x.get('name'))[:60])}</td><td>{e(str(x.get('customer'))[:40])}</td><td>{e(dl(x))}</td><td class=n>{(x.get('est_size_rub') or 0)/1e6:.1f}</td><td>{e(str(x.get('to_quote'))[:140])}</td></tr>")
h.append("</table><div class=g>Полные таблицы: Excel, листы «С НМЦ (рейтинг)» и «Без НМЦ (D)». Прибыль — после УСН 15%; медизделия без НДС, прочее с НДС 5% без вычета. Значения в тыс. ₽.</div></body></html>")
hp = f"{SCRATCH}/full1311.html"; open(hp, "w").write("".join(h))
subprocess.run(["/opt/pw-browsers/chromium", "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer", f"--print-to-pdf={os.path.abspath('../ZAKUPKI_1311_POLNYI_RAZBOR.pdf')}", "file://" + hp], check=True, capture_output=True, timeout=180)
print(len(R), cnt, round(tot_go), round(tot_c), len(top))
