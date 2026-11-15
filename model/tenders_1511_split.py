"""Две отдельные таблицы из tenders_1511_table.py: РФ — только новые закупки (поиск 15.11, ранее не отправлялись), Мир — все зарубежные.
→ ../TENDERY_RF_NOVYE_1511.xlsx/.pdf, ../TENDERY_MIR_1511.xlsx/.pdf"""
import json, os, html, subprocess
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from isp_letterhead import SCRATCH
R = json.load(open(f"{SCRATCH}/rows_1511.json"))
RU = json.load(open("data/tenders_1511/ru_chem.json"))["tenders"]
extra = {str(t.get("number")): t for t in RU}
H = PatternFill("solid", fgColor="0B2545"); HF = Font(bold=True, color="FFFFFF"); W = Alignment(wrap_text=True, vertical="top"); LINK = Font(color="0563C1", underline="single")
e = html.escape; f = lambda v: f"{v:,.0f}".replace(",", " ") if isinstance(v, (int, float)) else (e(str(v))[:20] if v else "—")

def build(name, rows, cols, getrow, linkcol, widths, title, pdfcols, pdfrow, note):
    wb = Workbook(); ws = wb.active; ws.title = "Тендеры"; ws.append(cols)
    for c in ws[1]: c.fill = H; c.font = HF; c.alignment = W
    for i, r in enumerate(rows, 1):
        ws.append(getrow(i, r))
        if r.get("link"): c = ws.cell(ws.max_row, linkcol); c.hyperlink = r["link"]; c.font = LINK
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = W
            if isinstance(c.value, (int, float)) and c.column > 1: c.number_format = "# ##0"
    for i, w in enumerate(widths): ws.column_dimensions[chr(65 + i)].width = w
    ws.freeze_panes = "C2"; ws.auto_filter.ref = ws.dimensions
    wb.save(f"../{name}.xlsx")
    css = ("@page{size:A4 landscape;margin:7mm}body{font-family:'DejaVu Sans',Arial;font-size:6.4px;color:#111}h1{font-size:13px;margin:0 0 2px;color:#0B2545}"
           "table{border-collapse:collapse;width:100%}th{background:#0B2545;color:#fff;padding:2px 3px;text-align:left}td{border-bottom:1px solid #ddd;padding:0 2px;vertical-align:top;line-height:1.15}.n{text-align:right;white-space:nowrap}.g{color:#555}a{color:#0563C1;text-decoration:none}")
    h = [f"<!DOCTYPE html><html><head><meta charset=utf-8><style>{css}</style></head><body><h1>{e(title)}</h1><div class=g>{e(note)}</div><table><tr>" + "".join(f"<th>{c}</th>" for c in pdfcols) + "</tr>"]
    for i, r in enumerate(rows, 1): h.append("<tr>" + "".join(pdfrow(i, r)) + "</tr>")
    h.append("</table></body></html>")
    hp = f"{SCRATCH}/{name}.html"; open(hp, "w").write("".join(h))
    subprocess.run(["/opt/pw-browsers/chromium", "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer", f"--print-to-pdf={os.path.abspath(f'../{name}.pdf')}", "file://" + hp], check=True, capture_output=True, timeout=180)
    import pypdf; print(name, len(rows), "rows,", len(pypdf.PdfReader(f"../{name}.pdf").pages), "pages")

def num(r): return f"<a href='{e(r['link'])}'>{e(r['n'])}</a>" if r.get("link") else e(r["n"])

# --- РФ: только новые (поиск 15.11)
ru = [r for r in R if r["src"] == "РФ (поиск 15.11)"]
for r in ru:
    t = extra.get(r["n"], {})
    r.update(law=t.get("law"), method=t.get("method"), nr=t.get("pp1875"), pos=(json.dumps(t.get("positions"), ensure_ascii=False)[:300] if t.get("positions") else ""), tz=("; ".join(t["tz_docs"][:2]) if isinstance(t.get("tz_docs"), list) and t.get("tz_docs") and isinstance(t["tz_docs"][0], str) else ""))
build("TENDERY_RF_NOVYE_1511", ru,
      ["№", "Номер", "Предмет", "Заказчик / регион", "Закон", "Способ", "Площадка", "НМЦ, ₽", "Подача (МСК)", "Нацрежим ПП 1875", "Что подходит", "Ссылка", "Позиции", "Сделано (✓)"],
      lambda i, r: [i, r["n"], r["subj"], r["cust"], r.get("law"), r.get("method"), r["plat"], r["nmc"], r["dl"], r.get("nr"), r["note"], r.get("link") or "—", r.get("pos"), ""],
      12, [4, 16, 46, 36, 8, 16, 16, 14, 14, 16, 28, 40, 50, 9],
      f"РФ: новые закупки химии (поиск 15.11), подача от 18.11 — {len(ru)} шт.",
      ["#", "Номер (ссылка)", "Предмет", "Заказчик", "НМЦ, ₽", "Подача", "ПП 1875", "Что подходит"],
      lambda i, r: [f"<td>{i}</td>", f"<td>{num(r)}</td>", f"<td>{e(r['subj'][:80])}</td>", f"<td>{e(r['cust'][:50])}</td>", f"<td class=n>{f(r['nmc'])}</td>", f"<td style='white-space:nowrap'>{e(str(r['dl'])[:16])}</td>", f"<td>{e(str(r.get('nr') or '—')[:22])}</td>", f"<td>{e(str(r['note'])[:50])}</td>"],
      "Ранее не отправлялись (выгрузка 13.11 — в ZAPROS_TZ_1411/TENDERY_TZ_1511). Подаёт Гидропас. Ограничение ПП 1875: иностранный товар отклоняется при ≥2 заявках с российским.")

# --- Мир
mir = [r for r in R if r["who"] == "ISP"]
order = {"Напрямую (возможно)": 0, "Уточнить": 1, "Через партнёра / местная регистрация": 2}
mir.sort(key=lambda r: (order.get(r["acc"], 3), r["d"]))
build("TENDERY_MIR_1511", mir,
      ["№", "Доступ для ISP", "Страна / заказчик", "Номер", "Предмет и объём", "Подача", "Оценка стоимости", "Площадка", "Ссылка", "Как получить документы", "Условия допуска", "Сделано (✓)"],
      lambda i, r: [i, r["acc"], r["cust"], r["n"], r["subj"], r["dl"], r["nmc"], r["plat"], r.get("link") or "—", r["todo"], r["note"], ""],
      9, [4, 18, 34, 20, 56, 16, 16, 16, 42, 40, 44, 9],
      f"Мир: тендеры для ISP Group, подача от 18.11 — {len(mir)} шт.",
      ["#", "Доступ", "Заказчик / страна", "Номер (ссылка)", "Предмет", "Подача"],
      lambda i, r: [f"<td>{i}</td>", f"<td>{e(r['acc'][:14])}</td>", f"<td>{e(r['cust'][:30])}</td>", f"<td>{num(r)[:400]}</td>", f"<td>{e(r['subj'][:58])}</td>", f"<td style='white-space:nowrap'>{e(str(r['dl'])[:16])}</td>"],
      "Сначала — куда можно подать напрямую, затем «уточнить», затем — через партнёра/местную регистрацию. Доступ определён автоматически по тексту условий, перед подачей проверить. Без Ирана, Сирии и OFAC SDN.")
