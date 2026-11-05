"""Маржа будущего оптовика (ОСНО) по 51 конкурсу ниши. data/companies/niche_margin.json → ../NISHA_MARZHA.pdf + .xlsx"""
import json, html, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from pdfutil import html_to_pdf
from tidy import tidy

d = json.load(open("data/companies/niche_margin.json")); T = d["tenders"]; S = d["summary"]
e = lambda s: html.escape(str(s if s is not None else "—"))
fm = lambda v: f"{v:,.0f}".replace(",", " ") if isinstance(v, (int, float)) else "—"
pc = lambda v: f"{v*100:.0f} %" if isinstance(v, (int, float)) else "—"
def J(v):
    if isinstance(v, dict): return ", ".join(f"{k}: {J(x)}" for k, x in v.items())
    if isinstance(v, list): return "; ".join(J(x) for x in v)
    return str(v if v is not None else "")
cost = lambda t: (t.get("cost") or {}).get("total_no_vat") if isinstance(t.get("cost"), dict) else t.get("cost")
M = sorted([t for t in T if t.get("margin_m10") is not None], key=lambda t: -((t.get("nmc_no_vat") or 0) * 0.9 - (cost(t) or 0)))
N = [t for t in T if t.get("margin_m10") is None]
css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:8.2px;color:#111;margin:0}h1{font-size:14px;margin:0 0 3px}h2{font-size:10.5px;margin:6px 0 3px;border-bottom:2px solid #305496}
table{border-collapse:collapse;width:100%;margin:2px 0 4px}th,td{border:1px solid #bbb;padding:1px 3px;vertical-align:top}th{background:#305496;color:#fff;text-align:left}td.n{text-align:right;white-space:nowrap}
.ok{background:#c6efce}.warn{background:#fff2cc}.bad{background:#f8cbad}.small{font-size:7px;color:#444}.box{border:1px solid #305496;padding:4px 6px;background:#f3f6fb;margin:4px 0}"""
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body><h1>Маржа будущей компании-оптовика (ОСНО) по конкурсам ниши — {len(T)} лотов</h1>",
     f"<div class='small'>{e(S.get('method'))} Курс ЦБ 26.10.2026; доставка — выделенная машина ≈95 ₽/км (консервативно). Документация закупок из сессии недоступна (zakupki.gov.ru, b2b-center) — база НДС принята по типу заказчика: {e(S.get('vat_rule'))[:260]}</div>",
     f"<div class='box'><b>Итог:</b> маржу удалось посчитать по {len(M)} лотам из {len(T)} — все {len(M)} дают ≥15 % даже при −10 % (энергетика: гликоль, ингибитор, активированные угли, гидразин, иониты). "
     "По остальным маржу не посчитать без КП: брендовые реагенты одного производителя (Экоэнерго «Оптион», СИНАЛАБ, Эмикор, Инкорамин — победители в прошлых контрактах почти не снижали цену, значит маржа есть только при дилерской скидке), "
     "реагенты нефтедобычи с обязательными испытаниями (Русвьетпетро), катализаторы НПЗ на закрытых площадках (цены и количества скрыты), запчасти/оборудование/услуги — не наш профиль. "
     "Вывод: стартовать с реагентов и материалов водоподготовки энергетики — там простой товар, открытые ЕИС-лоты и высокая маржа.</div>",
     "<h2>1. Лоты с посчитанной маржой (по прибыли при −10 %)</h2><table><tr><th>№</th><th>Заказчик</th><th>Предмет</th><th>НМЦ без НДС</th><th>Себест. без НДС</th><th>Маржа по НМЦ</th><th>При −10 %</th><th>Прибыль при −10 %</th><th>Увер.</th><th>Что проверить</th></tr>"]
for t in M:
    p = (t.get("nmc_no_vat") or 0) * 0.9 - (cost(t) or 0)
    h.append(f"<tr class='{'ok' if t['confidence'] != 'низкая' else 'warn'}'><td class=small>{e(t['number'])}</td><td class=small>{e(t['customer'])[:45]}</td><td>{e(t['subject'])[:80]}</td><td class=n>{fm(t.get('nmc_no_vat'))}</td><td class=n>{fm(cost(t))}</td>"
             f"<td class=n>{pc(t['margin_nmc'])}</td><td class=n>{pc(t['margin_m10'])}</td><td class=n>{fm(p)}</td><td>{e(t['confidence'])}</td><td class=small>{e(J(t.get('blockers')) or J(t.get('notes')))[:200]}</td></tr>")
h.append(f"<tr><th colspan=3>Итого {len(M)} лотов</th><th class=n>{fm(sum(t.get('nmc_no_vat') or 0 for t in M))}</th><th class=n>{fm(sum(cost(t) or 0 for t in M))}</th><th></th><th></th><th class=n>{fm(sum((t.get('nmc_no_vat') or 0)*.9-(cost(t) or 0) for t in M))}</th><th colspan=2></th></tr></table>")
h.append("<h2>2. Почему по остальным маржи нет / что запросить</h2><table><tr><th>№</th><th>Заказчик</th><th>Предмет</th><th>НМЦ</th><th>Причина / у кого КП</th></tr>")
for t in N:
    h.append(f"<tr><td class=small>{e(t['number'])}</td><td class=small>{e(t['customer'])[:40]}</td><td>{e(t['subject'])[:75]}</td><td class=n>{fm(t.get('nmc'))}</td><td class=small>{e(J(t.get('blockers')) or J(t.get('notes')))[:230]}</td></tr>")
h.append("</table></body></html>")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/nmarg.html"
open(hp, "w").write("".join(h)); html_to_pdf(hp, "../NISHA_MARZHA.pdf")

wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Маржа по лотам"
ws.append(["№", "Заказчик", "Предмет", "Сегмент", "НМЦ", "НМЦ без НДС", "База НДС", "Себестоимость без НДС", "Маржа по НМЦ", "Маржа при −10 %", "Уверенность", "Блокеры", "Примечания"])
for t in M + N:
    ws.append([t["number"], t["customer"], t["subject"], t.get("segment"), t.get("nmc"), t.get("nmc_no_vat"), J(t.get("nmc_vat_basis")), cost(t), t.get("margin_nmc"), t.get("margin_m10"), t.get("confidence"), J(t.get("blockers")), J(t.get("notes"))[:2000]])
    X = ws.max_row
    for j in (5, 6, 8): ws.cell(X, j).number_format = "#,##0"
    for j in (9, 10): ws.cell(X, j).number_format = "0%"
s2 = wb.create_sheet("Позиции"); s2.append(["№ лота", "Позиция", "Кол-во", "Ед.", "НМЦ ед.", "Рынок ед.", "Источник"])
for t in T:
    for it in t.get("items") or []:
        s2.append([t["number"], J(it.get("name")), it.get("qty"), J(it.get("unit")), it.get("nmc_unit"), it.get("market_unit"), J(it.get("source"))[:1000]])
s3 = wb.create_sheet("Кому запросить КП"); s3.append(["Запрос"]); [s3.append([J(x)]) for x in (S.get("kp_requests") or [])]
s4 = wb.create_sheet("Методика"); [s4.append([k, J(v)[:3000]]) for k, v in S.items() if k not in ("kp_requests",)]
for s in wb:
    for c in s[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
    for i in range(1, 14): s.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 26
s3.column_dimensions["A"].width = 140; s4.column_dimensions["B"].width = 140
tidy(wb, row_h=30); wb.save("../NISHA_MARZHA.xlsx")
print(len(M), len(N))
