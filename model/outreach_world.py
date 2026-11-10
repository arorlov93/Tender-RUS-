"""Рассылка ISP Group LLC по миру: покупатели ионообменных смол и фосфонатов/антискалантов (10.11.2026).
Данные: data/companies/outreach_world_central_asia.json, outreach_world_mena_asia.json (e-mail подтверждён в 2 источниках → verified).
→ ../ISP_RASSYLKA_SMOLY_FOSFONATY_1011.xlsx + .pdf"""
import json, html, openpyxl
from openpyxl.styles import Font, PatternFill
from pdfutil import html_to_pdf
from tidy import tidy
SUBJ = "Ion exchange resins and water treatment chemicals: supply offer from ISP Group LLC"
LETTER = open("data/companies/outreach_letter_en.txt").read()
R = []
for f, reg in (("outreach_world_central_asia.json", "Центр. Азия / Кавказ / Монголия"), ("outreach_world_mena_asia.json", "Турция / БВ / Африка / Азия")):
    for x in json.load(open("data/companies/" + f)):
        R.append(dict(reg=reg, c=x["company"], cc=x.get("country", ""), t=x.get("buyer_type", ""), p=x.get("products", ""),
                      e=x.get("procurement_email") or "", oe=x.get("other_emails") or "", ph=x.get("phone") or "",
                      portal=x.get("tender_portal") or x.get("supplier_registration_url") or "", v=bool(x.get("verified")),
                      src=" ; ".join(x.get("sources", [])[:2]), note=x.get("note", "")))
R.sort(key=lambda r: (not r["v"], r["reg"], r["cc"], r["c"]))
oe = lambda v: ", ".join(v) if isinstance(v, list) else str(v)
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Кому писать"
ws.append(["Регион", "Компания", "Страна", "Тип", "Что предлагать", "E-mail (основной)", "Проверен 2 источниками", "Другие e-mail", "Телефон", "Портал / регистрация поставщика", "Тема письма", "Источники", "Примечание"])
for r in R:
    ws.append([r["reg"], r["c"], r["cc"], r["t"], {"resins": "смолы", "phosphonates": "фосфонаты/антискаланты", "both": "смолы + фосфонаты"}.get(r["p"], r["p"]),
               r["e"], "да" if r["v"] else "нет (звонок/уточнить)", oe(r["oe"]), oe(r["ph"]), r["portal"], SUBJ, r["src"], r["note"]])
    ws.cell(ws.max_row, 7).fill = PatternFill("solid", fgColor="c6efce" if r["v"] else "fff2cc")
s2 = wb.create_sheet("Текст письма"); s2.append(["Тема", SUBJ]); s2.append(["Текст", LETTER])
s2.column_dimensions["A"].width = 10; s2.column_dimensions["B"].width = 120
for s in wb:
    for c in s[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")
for i, w in enumerate([20, 36, 12, 16, 18, 30, 12, 30, 20, 30, 30, 40, 50], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
tidy(wb, row_h=45); wb.save("../ISP_RASSYLKA_SMOLY_FOSFONATY_1011.xlsx")
e = html.escape
nv = sum(r["v"] for r in R)
css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:7.4px;margin:0}h1{font-size:12px;margin:0 0 3px}table{border-collapse:collapse;width:100%}
th,td{border:1px solid #bbb;padding:1.5px 3px;vertical-align:top}th{background:#305496;color:#fff;text-align:left}.ok{background:#c6efce}.w{background:#fff2cc}.box{border:1px solid #305496;padding:3px 6px;background:#f3f6fb;margin:2px 0 4px}"""
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body><h1>ISP Group LLC: кому предложить смолы и фосфонаты (10.11.2026)</h1>",
     f"<div class=box>{len(R)} компаний, у {nv} e-mail подтверждён двумя источниками (зелёные; по ним можно слать сразу), остальные — сначала звонок/проверка. "
     "Все названия сверены со списком OFAC SDN — совпадений нет. Первый шаг в Казахстане: регистрация на zakup.sk.kz (Казатомпром, Самрук-Энерго, КМГ разом). "
     "Тема письма у всех одна; текст — лист «Текст письма» в Excel.</div><table><tr><th>Компания</th><th>Страна</th><th>Тип</th><th>Что</th><th>E-mail</th><th>Телефон</th><th>Портал</th></tr>"]
for r in R:
    h.append(f"<tr><td>{e(r['c'][:60])}</td><td>{e(r['cc'])}</td><td>{e(r['t'][:22])}</td><td>{e(r['p'])}</td><td class={'ok' if r['v'] else 'w'}>{e(r['e'])}</td><td>{e(oe(r['ph'])[:30])}</td><td>{e(r['portal'][:45])}</td></tr>")
h.append("</table></body></html>")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/isp_out.html"
open(hp, "w").write("".join(h)); html_to_pdf(hp, "../ISP_RASSYLKA_SMOLY_FOSFONATY_1011.pdf")
print(len(R), nv)
