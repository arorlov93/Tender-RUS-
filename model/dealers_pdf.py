"""Дилерства СПб НХК и БПО Евразия-Снаб: проверка заявленных партнёров + где получить авторизацию. data/companies/dealers_out.json → ../DILERSTVA_NHK_EURASIA.pdf + .xlsx"""
import json, html, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from pdfutil import html_to_pdf
from tidy import tidy

d = json.load(open("data/companies/dealers_out.json"))
e = lambda s: html.escape(str(s if s is not None else "—"))
def T(v):
    if isinstance(v, dict): return ", ".join(f"{k}: {T(x)}" for k, x in v.items() if x not in (None, ""))
    if isinstance(v, list): return "; ".join(T(x) for x in v)
    return str(v if v is not None else "")
cls = lambda s: "bad" if str(s).startswith("опроверг") else "warn" if str(s).startswith("не подтв") else "ok"
css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:8.1px;color:#111;margin:0}h1{font-size:14px;margin:0 0 3px}
h2{font-size:10.5px;margin:6px 0 3px;border-bottom:2px solid #305496;padding-bottom:1px}table{border-collapse:collapse;width:100%;margin:2px 0 4px}
th,td{border:1px solid #bbb;padding:1px 3px;vertical-align:top}th{background:#305496;color:#fff;text-align:left}
.ok{background:#c6efce}.bad{background:#f8cbad}.warn{background:#fff2cc}.small{font-size:7px;color:#444}.box{border:1px solid #305496;padding:4px 6px;background:#f3f6fb;margin:4px 0}ol{margin:2px 0 2px 16px;padding:0}"""
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body><h1>Дилерства: ООО «СПб НХК» и ООО «БПО Евразия-Снаб» — проверка и где получить</h1>",
     "<div class='small'>Открытые источники на 05.11.2026 (сайты производителей, годовые отчёты, СМИ, трекеры ухода из РФ, торговые базы, конференции). Окончательно дилерство подтверждает только производитель (письмо/сертификат). Полный список источников — в Excel.</div>",
     "<div class='box'><b>Итог: ни одно действующее дилерство обеих компаний по открытым источникам не подтверждается.</b> СПб НХК: «авторизованный представитель Зульцер Хемтех» — Sulzer ушёл из РФ, ООО «Зульцер Хемтех» с 10.2023 = ООО «Сплитекс»; старый документ в тендер подавать нельзя, нужно новое письмо «Сплитекс». Napco и «китайские производители» — не подтверждены. "
     "БПО Евразия-Снаб: прежние партнёры (Johnson Matthey, Honeywell UOP, Evonik) из РФ ушли; Chevron Lummus — канал закрыт санкциями; Zhejiang Fenghong работает, но договора не видно — ниша открыта.</div>",
     "<h2>1. Проверка заявленных партнёров</h2><table><tr><th>Компания</th><th>Партнёр</th><th>Статус</th><th>Доказательства</th></tr>"]
for v in d["verify"]:
    h.append(f"<tr class='{cls(v.get('status'))}'><td>{e(v.get('company'))}</td><td>{e(v.get('partner'))[:90]}</td><td>{e(v.get('status'))[:60]}</td><td class=small>{e(v.get('evidence'))[:420]}</td></tr>")
h.append("</table><h2>2. Где получить дилерство / авторизацию</h2><table><tr><th>Сегмент</th><th>Производитель</th><th>Страна</th><th>Дилерская программа</th><th>Контакты</th><th>Кому</th></tr>")
for w in d["where_to_get"]:
    c = w.get("contacts") or {}
    h.append(f"<tr><td class=small>{e(w.get('segment'))[:60]}</td><td>{e(w.get('manufacturer'))[:80]}</td><td class=small>{e(w.get('country'))[:20]}</td><td class=small>{e(w.get('dealer_program'))[:160]}</td>"
             f"<td class=small>{e(c.get('site'))[:40]}<br>{e(c.get('phone'))[:50]}<br>{e(c.get('email'))[:40]}</td><td class=small>{e(w.get('fits'))[:60]}</td></tr>")
h.append("</table><h2>3. Что даёт дилерство в тендерах</h2><ul>" + "".join(f"<li>{e(T(x.get('point') if isinstance(x, dict) else x))[:300]}</li>" for x in (d.get("tender_value") or [])) + "</ul>")
h.append("<h2>4. Что и у кого запросить</h2><ol>" + "".join(f"<li>{e(T(x))[:600]}</li>" for x in d["how_to_request"]) + "</ol></body></html>")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/dealers.html"
open(hp, "w").write("".join(h)); html_to_pdf(hp, "../DILERSTVA_NHK_EURASIA.pdf")

wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Проверка партнёров"
cols = ["company", "partner", "claimed_by", "status", "evidence", "source", "date"]
ws.append(cols)
for v in d["verify"]: ws.append([T(v.get(c)) for c in cols])
s2 = wb.create_sheet("Где получить"); cols2 = ["segment", "manufacturer", "country", "inn", "dealer_program", "site", "phone", "email", "verified", "fits", "why", "sources"]
s2.append(cols2)
for w in d["where_to_get"]:
    c = w.get("contacts") or {}
    s2.append([T(w.get(k)) if k not in ("site", "phone", "email") else T(c.get(k)) for k in cols2])
s3 = wb.create_sheet("Что запросить"); s3.append(["Шаг"]); [s3.append([T(x)]) for x in d["how_to_request"]]
s4 = wb.create_sheet("Тендеры"); s4.append(["Пункт", "Источник"]); [s4.append([T(x.get("point")), T(x.get("source"))] if isinstance(x, dict) else [T(x)]) for x in (d.get("tender_value") or [])]
s5 = wb.create_sheet("Источники"); s5.append(["Источник"]); [s5.append([T(x)]) for x in d.get("sources", [])]
for s in wb:
    for c in s[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
    for i in range(1, 13): s.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 32
for s in (s3, s4, s5): s.column_dimensions["A"].width = 140
tidy(wb, row_h=40); wb.save("../DILERSTVA_NHK_EURASIA.xlsx")
print(len(d["verify"]), len(d["where_to_get"]), len(d["how_to_request"]))
