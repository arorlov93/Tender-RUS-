"""Итоговая таблица по запросам КП 13.11.2026: спецификация → закупки (ссылки, НМЦ, подача) → заводы (проверенные e-mail/телефоны).
→ ../ZAPROSY_KP_1311_ITOG.pdf + лист «0. Итог» в ../ZAPROSY_KP_1311.xlsx"""
import json, html, os, subprocess, re
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from isp_letterhead import SCRATCH

S = json.load(open("data/rfq_1311/specs.json"))
C = "data/companies/"
def lj(f):
    d = json.load(open(C + f)); return d if isinstance(d, list) else next(v for v in d.values() if isinstance(v, list))
VLA = lj("rfq_video_light_appl.json")
G = lambda k: [x for x in VLA if str(x.get("group", "")).endswith(k)]
FACT = {1: lj("rfq_cable_plants.json"), 2: lj("rfq_tube_makers.json"), 3: G("video"), 4: G("light"), 5: G("washers")}
SHORT = {1: "Кабель", 2: "Пробирки, контейнеры, лаб. расходники", 3: "Видеонаблюдение", 4: "Светильники ЕГП", 5: "Стиральные машины"}
# Сроки подачи, проверенные по карточкам ЕИС (endDT): подача уже прошла
PAST = {"0356100024826000029": "подача была 22.10", "0353100005326000014": "подача была 09.11"}

wb = load_workbook("../ZAPROSY_KP_1311.xlsx")
L1 = {}
for c in wb["1. Закупки и ссылки"].iter_rows(min_row=2):
    v = [x.value for x in c]
    L1[str(v[2])] = dict(status=v[1], subj=v[3], cust=v[4], etp=v[5], nmc=v[6], dl=v[7], term=v[8], url=c[9].hyperlink.target if c[9].hyperlink else None)

def dkey(d):
    m = re.match(r"(\d\d)\.(\d\d)", str(d or "")); return (int(m[2]), int(m[1])) if m else (99, 99)

rows = []  # (spec, lot dict)
summ = []
for s in S:
    lots = []
    for l in s["lots"]:
        i = L1.get(l["lot"], {})
        st = "Исключить: " + PAST[l["lot"]] if l["lot"] in PAST else (i.get("status") or "")
        lots.append(dict(lot=l["lot"], cust=l["customer"], nmc=i.get("nmc"), dl=i.get("dl") or l.get("deadline"), url=i.get("url"), st=st, npos=len(l["positions"]),
                         note="позиции — из документации ЕИС (скачать)" if not l["positions"] or l["lot"] == "32616558994" else ""))
    lots.sort(key=lambda x: dkey(x["dl"]))
    act = [x for x in lots if not x["st"].startswith("Исключить")]
    fac = [x for x in FACT[s["no"]] if not x.get("exclude") and x.get("company") and (x.get("email") or x.get("phone"))]
    ok = [x for x in fac if x.get("verified")]
    summ.append(dict(no=s["no"], what=SHORT[s["no"]], n=len(act), npos=sum(x["npos"] for x in act), nmc=sum(x["nmc"] or 0 for x in act),
                     first=min((x["dl"] for x in act), key=dkey, default="—"), ok=len(ok), tot=len(fac), lots=lots, fac=sorted(fac, key=lambda x: not x.get("verified"))))

# ---- Excel: лист «0. Итог»
if "0. Итог" in wb.sheetnames: del wb["0. Итог"]
ws = wb.create_sheet("0. Итог", 0)
H = PatternFill("solid", fgColor="0B2545"); HF = Font(bold=True, color="FFFFFF"); W = Alignment(wrap_text=True, vertical="top")
SUB = PatternFill("solid", fgColor="DCE6F2"); LINK = Font(color="0563C1", underline="single"); RED = Font(color="B00020", bold=True)
def hdr(vals):
    ws.append(vals)
    for c in ws[ws.max_row]: c.fill = H; c.font = HF; c.alignment = W
hdr(["Спец. №", "Что запрашиваем", "Закупок (актуальных)", "Позиций", "Сумма НМЦ, ₽", "Ближайшая подача", "Заводов проверено / всего", "Вложение"])
for s in summ:
    ws.append([s["no"], s["what"], s["n"], s["npos"], s["nmc"], s["first"], f"{s['ok']} / {s['tot']}", f"Spec_{s['no']}.xlsx"])
ws.append(["Итого", "", sum(s["n"] for s in summ), sum(s["npos"] for s in summ), sum(s["nmc"] for s in summ), "", f"{sum(s['ok'] for s in summ)} / {sum(s['tot'] for s in summ)}", ""])
for c in ws[ws.max_row]: c.font = Font(bold=True)
for s in summ:
    ws.append([]); ws.append([f"Спецификация № {s['no']}: {s['what']}"]); ws.cell(ws.max_row, 1).font = Font(bold=True, size=12)
    hdr(["", "Закупка", "Заказчик", "НМЦ, ₽", "Подача", "Статус", "Ссылка", "Примечание"])
    for x in s["lots"]:
        ws.append(["", x["lot"], x["cust"], x["nmc"], x["dl"], x["st"], x["url"] or "—", x["note"]])
        if x["url"]: ws.cell(ws.max_row, 7).hyperlink = x["url"]; ws.cell(ws.max_row, 7).font = LINK
        if x["st"].startswith("Исключить"): ws.cell(ws.max_row, 6).font = RED
    hdr(["", "Завод", "ИНН", "E-mail", "Телефон", "Проверено", "Сайт", "Примечание"])
    for x in s["fac"]:
        ws.append(["", x.get("company"), x.get("inn"), x.get("email"), x.get("phone"), "ДА" if x.get("verified") else "НЕТ (в письмо не включён)", x.get("site"), (x.get("notes") or "")[:200]])
        if not x.get("verified"): ws.cell(ws.max_row, 6).font = RED
for r in ws.iter_rows(min_row=2):
    for c in r: c.alignment = W
for col, w in zip("ABCDEFGH", [8, 34, 40, 26, 20, 22, 40, 50]): ws.column_dimensions[col].width = w
for r in ws.iter_rows():
    if isinstance(r[4].value, (int, float)) and r[4].column_letter == "E" and r[0].value != "": r[4].number_format = "# ##0"
    if isinstance(r[3].value, (int, float)) and r[0].value == "": r[3].number_format = "# ##0"
ws.freeze_panes = "A2"
wb.save("../ZAPROSY_KP_1311.xlsx")

# ---- PDF
e = html.escape; f = lambda v: f"{v:,.0f}".replace(",", " ") if isinstance(v, (int, float)) and v else "—"
css = ("@page{size:A4 landscape;margin:7mm}body{font-family:'DejaVu Sans',Arial;font-size:7.2px;color:#111}h1{font-size:13px;margin:0 0 2px;color:#0B2545}"
       "h2{font-size:9px;margin:5px 0 1px;color:#0B2545}table{border-collapse:collapse;width:100%}th{background:#0B2545;color:#fff;padding:1px 3px;text-align:left}"
       "td{border-bottom:1px solid #ddd;padding:1px 3px;vertical-align:top}.n{text-align:right;white-space:nowrap}.g{color:#555}.r{color:#B00020;font-weight:bold}"
       ".cols{display:flex;gap:6px}.cols>div{flex:1}a{color:#0563C1;text-decoration:none}")
h = [f"<!DOCTYPE html><html><head><meta charset=utf-8><style>{css}</style></head><body><h1>Запросы КП 13.11.2026 — итоговая таблица</h1>",
     "<table><tr><th>Спец.</th><th>Что запрашиваем</th><th>Закупок</th><th>Позиций</th><th>Сумма НМЦ, ₽</th><th>Ближайшая подача</th><th>Заводов: проверено / всего</th><th>Вложение</th></tr>"]
for s in summ:
    h.append(f"<tr><td>{s['no']}</td><td>{e(s['what'])}</td><td class=n>{s['n']}</td><td class=n>{s['npos']}</td><td class=n>{f(s['nmc'])}</td><td>{e(str(s['first']))}</td><td class=n>{s['ok']} / {s['tot']}</td><td>Spec_{s['no']}.xlsx</td></tr>")
h.append(f"<tr><td colspan=2><b>Итого</b></td><td class=n><b>{sum(s['n'] for s in summ)}</b></td><td class=n><b>{sum(s['npos'] for s in summ)}</b></td><td class=n><b>{f(sum(s['nmc'] for s in summ))}</b></td><td></td><td class=n><b>{sum(s['ok'] for s in summ)} / {sum(s['tot'] for s in summ)}</b></td><td></td></tr></table>")
for s in summ:
    h.append(f"<h2>Спецификация № {s['no']}: {e(s['what'])}</h2><div class=cols><div><table><tr><th>Закупка (ссылка)</th><th>Заказчик</th><th>НМЦ, ₽</th><th>Подача</th><th>Статус</th></tr>")
    for x in s["lots"]:
        num = f"<a href='{e(x['url'])}'>{e(x['lot'])}</a>" if x["url"] else e(x["lot"])
        h.append(f"<tr><td>{num}</td><td>{e(x['cust'][:48])}{' <i class=g>(' + e(x['note']) + ')</i>' if x['note'] else ''}</td><td class=n>{f(x['nmc'])}</td><td>{e(str(x['dl'] or ''))}</td><td class={'r' if x['st'].startswith('Исключить') else ''}>{e(x['st'])}</td></tr>")
    h.append("</table></div><div><table><tr><th>Завод</th><th>E-mail</th><th>Телефон</th><th>Пров.</th></tr>")
    for x in s["fac"]:
        nm = re.sub(r"\s*\(.*?\)", "", x.get("company") or "")[:40]
        h.append(f"<tr><td>{e(nm)}</td><td>{e(x.get('email') or '—')}</td><td style='white-space:nowrap'>{e(x.get('phone') or '—')}</td><td class={'' if x.get('verified') else 'r'}>{'да' if x.get('verified') else 'нет'}</td></tr>")
    if not s["fac"]: h.append("<tr><td colspan=4 class=r>Проверенных контактов нет</td></tr>")
    h.append("</table></div></div>")
h.append("<div class=g style='margin-top:4px'>«Пров.» — e-mail и телефон совпали в 2 независимых источниках; в письма идут только проверенные. Исключены: Камкабель (OFAC по OpenSanctions, проверить), Кронт и Медполимер (не делают пробирки). "
         "Две закупки видео исключены: подача по ЕИС уже прошла. Письма отправляются с info@gidropas.ru, тексты — лист «4. Письма». НМЦ по B2B — оценка из выгрузки.</div></body></html>")
hp = f"{SCRATCH}/itog1311.html"; open(hp, "w").write("".join(h))
subprocess.run(["/opt/pw-browsers/chromium", "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer", f"--print-to-pdf={os.path.abspath('../ZAPROSY_KP_1311_ITOG.pdf')}", "file://" + hp], check=True, capture_output=True, timeout=180)
for s in summ: print(s["no"], s["what"], s["n"], s["npos"], round(s["nmc"]), s["first"], s["ok"], s["tot"])
