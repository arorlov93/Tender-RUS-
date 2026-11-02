"""Список для тендерщицы: только выгодные закупки с подачей с 06.11 — номер, ссылка, площадка, срок. Без наших цен и маржи. → ../ZAKUPKI_0211_TZ_SKACHAT.xlsx"""
import json, runpy, io, contextlib, openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from tidy import tidy
with contextlib.redirect_stdout(io.StringIO()):
    g = runpy.run_path("zakupki_0211.py")
rows = g["rows"]
V = [x for x in json.load(open("data/zakupki_0211/verdict.json")) if x[1].startswith("ВЫГОДНО") and str(rows[x[0]]["Окончание приема заявок"] or "") >= "2026-11-06"]
V.sort(key=lambda x: str(rows[x[0]]["Окончание приема заявок"]))
def link(n, r):
    u = str(r.get("Ссылка на ЕИС") or "").strip()
    if u.startswith("http"): return u
    if n.isdigit() and len(n) == 11 and n.startswith("326"): return f"https://zakupki.gov.ru/epz/order/notice/notice223/common-info.html?regNumber={n}"
    if n.isdigit() and len(n) == 19: return f"https://zakupki.gov.ru/epz/order/extendedsearch/results.html?searchString={n}"
    return ""
thin = Side(style="thin", color="BFBFBF"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Скачать ТЗ"
ws.append(["Закупки для подачи (с 06.11): скачать документацию — ТЗ/спецификацию, проект договора, обоснование НМЦ"]); ws["A1"].font = Font(bold=True, size=12)
ws.append(["№", "Номер закупки", "Ссылка", "Площадка (ЭТП)", "Подача до", "Заказчик", "Предмет", "Отметка: ТЗ скачано"])
for c in ws[2]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
for i, x in enumerate(V, 1):
    n = x[0]; r = rows[n]; u = link(n, r)
    ws.append([i, n, u or "искать по номеру на площадке", r["ЭТП"], str(r["Окончание приема заявок"] or "")[:16], r["Заказчик"], r["Название"], ""])
    X = ws.max_row
    if u: ws.cell(X, 3).hyperlink = u; ws.cell(X, 3).font = Font(color="0563C1", underline="single")
    for c in ws[X]: c.border = BOX; c.alignment = Alignment(wrap_text=True, vertical="top")
for i, w in enumerate([4, 22, 45, 20, 16, 34, 60, 14], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "C3"
tidy(wb, row_h=36); wb.save("../ZAKUPKI_0211_TZ_SKACHAT.xlsx")
for x in V: print(x[0], str(rows[x[0]]["Окончание приема заявок"])[:16], rows[x[0]]["ЭТП"], "|", link(x[0], rows[x[0]])[:90])
