"""Таблица «Где запросить ТЗ» (14.11.2026): только лоты со статусом «Запросить ТЗ», срок исполнения ≤6 мес, со ссылками.
→ ../ZAPROS_TZ_1411.xlsx + .pdf"""
import re, html, os, subprocess
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from isp_letterhead import SCRATCH

src = load_workbook("../ZAPROSY_KP_1311.xlsx")["1. Закупки и ссылки"]
PLAT = {"АСТ ГОЗ": "https://www.astgoz.ru/", "ЭТП ГПБ": "https://etpgpb.ru/", "РТС-тендер": "https://www.rts-tender.ru/"}
def eis(n):
    if re.fullmatch(r"\d{11}", n): return f"https://zakupki.gov.ru/epz/order/notice/notice223/common-info.html?regNumber={n}"
    if re.fullmatch(r"\d{19}", n): return f"https://zakupki.gov.ru/epz/order/extendedsearch/results.html?searchString={n}"
def dkey(d):
    m = re.match(r"(\d\d)\.(\d\d) ?(\d\d:\d\d)?", str(d or "")); return (m[2], m[1], m[3] or "") if m else ("99",)
rows, skip = [], []
for r in src.iter_rows(min_row=2, values_only=True):
    if r[1] != "Запросить ТЗ": continue
    raw = str(r[2]); n = raw.split("/")[-1]; plat = str(r[5])
    closed = "АСТ ГОЗ" in plat
    b2b = f"https://www.b2b-center.ru/market/view.html?id={n}" if re.fullmatch(r"\d{7}", n) and "b2b" in plat.lower() else None
    link = b2b or (PLAT["АСТ ГОЗ"] if closed else None) or eis(n) or next((u for k, u in PLAT.items() if k in plat), None)
    how = ("Закрытая закупка 44-ФЗ: запрос документации через АСТ ГОЗ, срок — 3 рабочих дня (до ~17.11)" if closed else
           "Войти на B2B-Center → скачать документацию лота" if b2b else
           "ЕИС → вкладка «Документы» → скачать ТЗ/спецификацию" if eis(n) else
           "Найти по номеру на площадке " + plat + (" (нужна аккредитация; ссылки на площадку в выгрузке нет)" if "Кортрос" in plat else ""))
    rec = dict(n=n, subj=str(r[3]), cust=str(r[4]), plat=plat, nmc=r[6], dl=str(r[7] or ""), term=r[8], link=link, how=how, todo=str(r[11] or ""))
    (skip if isinstance(r[8], (int, float)) and r[8] > 6 else rows).append(rec)
rows.sort(key=lambda x: dkey(x["dl"]))

wb = Workbook(); ws = wb.active; ws.title = "Запросить ТЗ"
H = PatternFill("solid", fgColor="0B2545"); HF = Font(bold=True, color="FFFFFF"); W = Alignment(wrap_text=True, vertical="top")
URG = PatternFill("solid", fgColor="FDE2E1"); LINK = Font(color="0563C1", underline="single")
cols = ["№", "Номер закупки", "Предмет", "Заказчик / регион", "Площадка", "НМЦ, ₽", "Подача", "Ссылка (кликнуть)", "Как получить ТЗ", "Что дальше", "ТЗ скачано (✓)"]
ws.append(cols)
for c in ws[1]: c.fill = H; c.font = HF; c.alignment = W
for i, x in enumerate(rows, 1):
    ws.append([i, x["n"], x["subj"], x["cust"], x["plat"], x["nmc"], x["dl"], x["link"] or "искать по номеру на площадке", x["how"], x["todo"], ""])
    if x["link"]: c = ws.cell(ws.max_row, 8); c.hyperlink = x["link"]; c.font = LINK
    ws.cell(ws.max_row, 6).number_format = "# ##0"
    if x["dl"].startswith(("16.11", "17.11")):
        for c in ws[ws.max_row][:7]: c.fill = URG
for row in ws.iter_rows(min_row=2):
    for c in row: c.alignment = W
for col, w in zip("ABCDEFGHIJK", [4, 22, 40, 34, 18, 13, 12, 45, 40, 45, 10]): ws.column_dimensions[col].width = w
ws.freeze_panes = "C2"; ws.auto_filter.ref = ws.dimensions
w2 = wb.create_sheet("Не по фильтру (>6 мес)")
w2.append(["Номер", "Предмет", "Заказчик", "Срок, мес", "Подача"])
for x in skip: w2.append([x["n"], x["subj"], x["cust"], x["term"], x["dl"]])
wb.save("../ZAPROS_TZ_1411.xlsx")

e = html.escape; f = lambda v: f"{v:,.0f}".replace(",", " ") if isinstance(v, (int, float)) else "—"
css = ("@page{size:A4 landscape;margin:7mm}body{font-family:'DejaVu Sans',Arial;font-size:7.4px;color:#111}h1{font-size:13px;margin:0 0 2px;color:#0B2545}"
       "table{border-collapse:collapse;width:100%}th{background:#0B2545;color:#fff;padding:2px 3px;text-align:left}td{border-bottom:1px solid #ddd;padding:2px 3px;vertical-align:top}"
       ".n{text-align:right;white-space:nowrap}.u td{background:#FDE2E1}.g{color:#555}a{color:#0563C1;text-decoration:none}")
tot = sum(x["nmc"] or 0 for x in rows)
h = [f"<!DOCTYPE html><html><head><meta charset=utf-8><style>{css}</style></head><body><h1>Где запросить ТЗ — {len(rows)} закупок, НМЦ {tot/1e6:.1f} млн ₽</h1>",
     "<div class=g>Срок исполнения ≤6 мес. Розовым — подача 16–17.11: ТЗ скачать сегодня-завтра. Номер закупки — ссылка.</div>",
     "<table><tr><th>#</th><th>Номер (ссылка)</th><th>Предмет</th><th>Заказчик</th><th>Площадка</th><th>НМЦ, ₽</th><th>Подача</th><th>Как получить ТЗ</th></tr>"]
for i, x in enumerate(rows, 1):
    num = f"<a href='{e(x['link'])}'>{e(x['n'])}</a>" if x["link"] else e(x["n"])
    h.append(f"<tr class={'u' if x['dl'].startswith(('16.11','17.11')) else ''}><td>{i}</td><td>{num}</td><td>{e(x['subj'][:60])}</td><td>{e(x['cust'][:45])}</td><td>{e(x['plat'][:22])}</td>"
             f"<td class=n>{f(x['nmc'])}</td><td style='white-space:nowrap'>{e(x['dl'])}</td><td>{e(x['how'][:90])}</td></tr>")
h.append(f"</table><div class=g style='margin-top:4px'>Не по фильтру срока (>6 мес): {', '.join(x['n'] + ' (' + str(x['term']) + ' мес)' for x in skip)}. Ссылки на ЕИС по 44-ФЗ ведут на поиск по номеру. "
         "Источник: ZAKUPKI_1311_PODACHA_I_TZ.</div></body></html>")
hp = f"{SCRATCH}/tz1411.html"; open(hp, "w").write("".join(h))
subprocess.run(["/opt/pw-browsers/chromium", "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer", f"--print-to-pdf={os.path.abspath('../ZAPROS_TZ_1411.pdf')}", "file://" + hp], check=True, capture_output=True, timeout=180)
print(len(rows), round(tot), [(x["n"], x["term"]) for x in skip], sum(1 for x in rows if x["link"]), sum(1 for x in rows if x["dl"].startswith(("16.11", "17.11"))))
for x in rows:
    if not x["link"]: print("NO LINK", x["n"], x["plat"])
