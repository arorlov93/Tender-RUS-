"""Досье компании → PDF + Excel. Использование: python3 company_dossier.py data/companies/eurasia_out.json ../DOSIE_EURASIA_SNAB
Вход — JSON агента (company, profile, track_record, target_customers, current_tenders, sources, open_questions)."""
import json, html, sys, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from pdfutil import html_to_pdf
from tidy import tidy

src, out = sys.argv[1], sys.argv[2]
d = json.load(open(src)); C = d["company"]; P = d["profile"]
e = lambda s: html.escape(str(s if s is not None else "—"))
def T(v):
    if isinstance(v, dict): return ", ".join(f"{k}: {T(x)}" for k, x in v.items() if x not in (None, ""))
    if isinstance(v, list): return "; ".join(T(x) for x in v)
    return str(v if v is not None else "")
fm = lambda v: f"{v:,.0f}".replace(",", " ") if isinstance(v, (int, float)) else e(v)
css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:8.2px;color:#111;margin:0}h1{font-size:14px;margin:0 0 3px}
h2{font-size:10.5px;margin:6px 0 3px;border-bottom:2px solid #305496;padding-bottom:1px}table{border-collapse:collapse;width:100%;margin:2px 0 4px}
th,td{border:1px solid #bbb;padding:1px 3px;vertical-align:top}th{background:#305496;color:#fff;text-align:left}td.n{text-align:right;white-space:nowrap}
.small{font-size:7.1px;color:#444}.box{border:1px solid #305496;padding:4px 6px;background:#f3f6fb;margin:4px 0}ul{margin:2px 0 2px 14px;padding:0}"""
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body><h1>Досье: {e(C.get('name'))}</h1>",
     f"<div class='small'>ИНН {e(C.get('inn'))}, ОГРН {e(C.get('ogrn'))}. Источники — открытые (сайт, saby.ru, арбитражные дела, Госзатраты, gosplan/ЕИС, bicotender), на 05.11.2026. Гипотезы помечены. Полные данные и ссылки — в Excel.</div>",
     f"<div class='box'><b>Чем занимается:</b> {e(P.get('what_they_do'))}</div><h2>1. Компания</h2><table>"]
keys = [("Регистрация", "registered"), ("Адрес", "address"), ("Руководитель", "ceo"), ("Руководитель", "director"), ("Владельцы", "owners"), ("Связанные лица", "related_parties"),
        ("Численность", "employees"), ("ОКВЭД", "okved_main"), ("МСП", "msp"), ("Суды", "courts"), ("Лицензии", "licenses"), ("Налоги", "taxes_2025_rub"), ("Имущество", "property"), ("Рейтинг", "rating"), ("Контакты", "contacts"), ("Сайт", "site")]
for lab, k in keys:
    if C.get(k) not in (None, "", [], {}): h.append(f"<tr><th style='width:11%'>{lab}</th><td>{e(T(C[k]))[:420]}</td></tr>")
h.append("</table>")
fin = C.get("finance_mln_rub") or C.get("revenue_mln_rub")
if isinstance(fin, dict):
    h.append("<table><tr><th>Год</th>" + "".join(f"<th>{y}</th>" for y in fin) + "</tr><tr><td>Выручка, млн</td>" + "".join(f"<td class=n>{(v[0] if isinstance(v, list) else v)}</td>" for v in fin.values()) + "</tr>")
    if all(isinstance(v, list) for v in fin.values()): h.append("<tr><td>Чистая прибыль, млн</td>" + "".join(f"<td class=n>{v[1]}</td>" for v in fin.values()) + "</tr>")
    h.append("</table>")
h.append("<h2>2. Что может поставлять</h2><table><tr><th>Товарная группа</th><th>ОКПД2</th><th>Бренды/производители</th><th>Подтверждение</th></tr>")
for g in P.get("product_groups", []):
    h.append(f"<tr><td>{e(g.get('group'))}</td><td class=small>{e(g.get('okpd2'))}</td><td class=small>{e(T(g.get('brands')))[:160]}</td><td class=small>{e(g.get('evidence'))[:200]}</td></tr>")
h.append("</table><table><tr><th style='width:25%'>Отрасли</th><td>" + e(T(P.get("industries"))) + "</td></tr><tr><th>Регионы</th><td>" + e(T(P.get("regions"))) + "</td></tr>"
         "<tr><th>Сильные стороны</th><td><ul>" + "".join(f"<li>{e(x)}</li>" for x in P.get("strengths", [])) + "</ul></td></tr><tr><th>Слабые стороны</th><td><ul>" + "".join(f"<li>{e(x)}</li>" for x in P.get("weaknesses", [])) + "</ul></td></tr></table>")
h.append("<h2>3. История поставок (госзакупки и подтверждённые сделки)</h2><table><tr><th>Заказчик</th><th>Предмет</th><th>Сумма</th><th>Дата</th><th>№</th></tr>")
for t in d.get("track_record", []): h.append(f"<tr><td>{e(t.get('customer'))}</td><td>{e(t.get('subject'))[:150]}</td><td class=n>{fm(t.get('amount'))}</td><td>{e(t.get('date'))}</td><td class=small>{e(t.get('number'))[:60]}</td></tr>")
h.append("</table><h2>4. Целевые заказчики</h2><table><tr><th>Заказчик</th><th>ИНН</th><th>Отрасль</th><th>Почему / примеры закупок</th><th>Площадка</th></tr>")
for t in d.get("target_customers", []):
    h.append(f"<tr><td>{e(t.get('name'))}</td><td>{e(t.get('inn'))}</td><td class=small>{e(t.get('industry'))}</td><td class=small>{e(t.get('why'))[:170]}<br>{e(T(t.get('example_purchases')))[:230]}</td><td class=small>{e(t.get('platform'))[:60]}</td></tr>")
h.append("</table><h2>5. Текущие закупки под компанию (подача с 06.11.2026)</h2><table><tr><th>№</th><th>Заказчик</th><th>Предмет</th><th>НМЦ</th><th>Подача до</th><th>Соответствие</th></tr>")
for t in sorted(d.get("current_tenders", []), key=lambda t: str(t.get("deadline"))):
    h.append(f"<tr><td class=small>{e(t.get('number'))}</td><td class=small>{e(t.get('customer'))[:70]}</td><td>{e(t.get('subject'))[:120]}</td><td class=n>{fm(t.get('nmc'))}</td><td>{e(t.get('deadline'))[:16]}</td><td class=small>{e(t.get('fit'))[:110]}</td></tr>")
h.append("</table><h2>6. Открытые вопросы (проверить перед работой)</h2><ul>" + "".join(f"<li>{e(T(x))[:250]}</li>" for x in d.get("open_questions", [])) + "</ul></body></html>")
hp = f"/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/{out.split('/')[-1]}.html"
open(hp, "w").write("".join(h)); html_to_pdf(hp, out + ".pdf")

wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Компания"
ws.append(["Поле", "Значение"])
for k, v in C.items(): ws.append([k, T(v)[:3000]])
ws.append([]); ws.append(["Чем занимается", P.get("what_they_do")])
for k in ("industries", "regions", "strengths", "weaknesses"): ws.append([k, T(P.get(k))])
def sheet(name, rows, cols):
    s = wb.create_sheet(name); s.append(cols)
    for r in rows: s.append([T(r.get(c))[:3000] if not isinstance(r.get(c), (int, float)) else r.get(c) for c in cols])
    for c in s[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
    for i in range(1, len(cols) + 1): s.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 30
    return s
sheet("Товарные группы", P.get("product_groups", []), ["group", "okpd2", "brands", "evidence"])
sheet("История поставок", d.get("track_record", []), list((d.get("track_record") or [{}])[0].keys()) or ["customer"])
sheet("Целевые заказчики", d.get("target_customers", []), ["name", "inn", "industry", "why", "example_purchases", "platform"])
s = sheet("Текущие закупки", d.get("current_tenders", []), ["number", "customer", "subject", "nmc", "deadline", "url", "fit"])
for row in s.iter_rows(min_row=2):
    if row[5].value: row[5].hyperlink = row[5].value
    row[3].number_format = "#,##0"
s = wb.create_sheet("Источники и вопросы"); s.append(["Источники"]); [s.append([T(x)]) for x in d.get("sources", [])]; s.append([]); s.append(["Открытые вопросы"]); [s.append([T(x)]) for x in d.get("open_questions", [])]
s.column_dimensions["A"].width = 140
for c in wb["Компания"][1]: c.font = Font(bold=True)
wb["Компания"].column_dimensions["A"].width = 24; wb["Компания"].column_dimensions["B"].width = 140
tidy(wb, row_h=30); wb.save(out + ".xlsx")
print("ok", out)
