"""Маржа по РФ (Гидропас, УСН) на 15.11.2026: 5 разобранных ТЗ + закупки из разбора 13.11 с подачей ≥18.11 → ../RF_MARZHA_1511.xlsx/.pdf"""
import json, re, os, html, subprocess
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as CL
from isp_letterhead import SCRATCH
Z = "data/zakupki_1511_tz/"
A = [("32616549223", "Кузбассэнерго-РЭС", "Видеорегистраторы Страж-ПВР-05, 17 шт", "18.11 10:00", "kuzbass_32616549223", "Подавать"),
     ("0361200015026004877", "ЮСГКБ (Сахалин), 44-ФЗ", "Пробирки вакуумные 4 мл, 46 150 шт", "18.11 01:00 МСК", "sakh_0361200015026004877", "При условиях"),
     ("32616563695", "Хакасэнерго", "ОПС/видео, 7 позиций", "19.11 11:00", "hakas_32616563695", "При условиях"),
     ("32616573584", "Кировэнерго", "Расходники СДТУ, 73 позиции", "20.11 11:00", "kirov_32616573584", "Скорее нет"),
     ("32616554083", "Россети Томск", "Каталог 526 позиций (лимит)", "19.11 11:00", "tomsk_32616554083", "Не подавать")]
def g(d, *p):
    for k in p: d = d.get(k) if isinstance(d, dict) else None
    return d
R1 = []
for n, cust, what, dl, d, v in A:
    e = json.load(open(Z + d + "/analysis.json"))["economics"]
    sc = e.get("scenarios") or {}
    v5 = g(sc, "m10", "vat5", "net_profit") or g(e, "scenario_vat5_inside", "m10", "profit")
    if v5 is None and isinstance(sc, dict):
        for k2, val in sc.items():
            if "vat5" in k2 and isinstance(val, dict): v5 = g(val, "k0.9", "profit") or g(val, "m10", "profit"); break
    R1.append([n, cust, what, dl, e.get("revenue_nmc"), e.get("cost"), e.get("profit_nmc"), e.get("profit_m10"), e.get("margin_pct_m10"), e.get("profit_m20"), v5, e.get("breakeven_discount_pct"), v])
# 13.11
def dk(s):
    m = re.search(r"(\d\d)\.(\d\d)", str(s or "")); return (int(m[2]), int(m[1])) if m else None
src = load_workbook("../ZAKUPKI_1311_POLNYI_RAZBOR.xlsx").active
R2 = []
done = {r[0] for r in R1}
for r in src.iter_rows(min_row=2, values_only=True):
    n, dl = str(r[1]), r[10]; d = dk(dl)
    if not d or d < (11, 18) or n in done: continue
    if not isinstance(r[12], (int, float)): continue
    term = r[18]
    if isinstance(term, (int, float)) and term > 6: continue
    R2.append([n, r[3], r[5], r[4], r[7], r[9], dl, r[11], r[12], r[14], r[13], r[16], term])
R2.sort(key=lambda x: -(x[8] or 0))
H = PatternFill("solid", fgColor="0B2545"); HF = Font(bold=True, color="FFFFFF"); W = Alignment(wrap_text=True, vertical="top"); LINK = Font(color="0563C1", underline="single")
def link(n): return f"https://zakupki.gov.ru/epz/order/notice/notice223/common-info.html?regNumber={n}" if re.fullmatch(r"\d{11}", n) else f"https://zakupki.gov.ru/epz/order/extendedsearch/results.html?searchString={n}"
wb = Workbook()
def sheet(ws, title, cols, rows, widths, numcols, pctcols=()):
    ws.title = title; ws.append(cols)
    for c in ws[1]: c.fill = H; c.font = HF; c.alignment = W
    for r in rows:
        ws.append(r); c = ws.cell(ws.max_row, 1)
        if re.fullmatch(r"\d{11}|\d{19}", str(r[0])): c.hyperlink = link(str(r[0])); c.font = LINK
        for k in numcols: ws.cell(ws.max_row, k).number_format = "# ##0"
    for row in ws.iter_rows(min_row=2):
        for c in row: c.alignment = W
    for i, w in enumerate(widths, 1): ws.column_dimensions[CL(i)].width = w
    ws.freeze_panes = "B2"; ws.auto_filter.ref = ws.dimensions
sheet(wb.active, "Разбор ТЗ 15.11", ["Номер (ссылка)", "Заказчик", "Предмет", "Подача", "НМЦ (база), ₽", "Себестоимость, ₽", "Прибыль при НМЦ", "Прибыль −10%", "Маржа −10%, %", "Прибыль −20%", "Прибыль −10%, если НДС 5% внутри", "Безубыточное снижение, %", "Решение"],
      R1, [22, 24, 34, 16, 14, 14, 13, 13, 9, 13, 14, 10, 14], (5, 6, 7, 8, 10, 11))
sheet(wb.create_sheet(), "Разбор 13.11 подача ≥18.11", ["Номер (ссылка)", "Предмет", "Заказчик", "Регион", "Способ", "НМЦ, ₽", "Подача", "Прибыль при НМЦ", "Прибыль −10%", "Маржа −10%, %", "Прибыль −20%", "Решение 13.11", "Срок исп., мес"],
      R2, [22, 44, 34, 18, 16, 13, 13, 13, 13, 9, 13, 12, 8], (6, 8, 9, 11))
w3 = wb.create_sheet("Как считано")
for x in ["Прибыль — после УСН 15% (доходы минус расходы, минимум 1% от дохода), без НДС, если не указано иное.",
          "Лист 1: по присланным ТЗ — все позиции, цены поставщиков из открытых прайсов со ссылками (см. RF_RAZBOR_TZ_1511.xlsx), доставка — оценка.",
          "Вариант «НДС 5% внутри» — если доход Гидропаса за 2025 > 20 млн ₽ и заказчик примет цену как включающую НДС 5%.",
          "Лист 2: закупки из выгрузок 13.11 (полный разбор), подача 18.11 и позже, срок исполнения ≤6 мес; медизделия без НДС, прочее с НДС 5% без вычета; себестоимость пробирок — прайс BERIMED/МиниМед.",
          "ВНИМАНИЕ: в разборе 13.11 пробирки считались по МиниМед — его реестровые записи на вакуумные пробирки истекли 04.09.2026; перед подачей пересчитать по Еламед/Беримед."]: w3.append([x])
w3.column_dimensions["A"].width = 150
wb.save("../RF_MARZHA_1511.xlsx")
e = html.escape; k = lambda v: "—" if not isinstance(v, (int, float)) else f"{v:,.0f}".replace(",", " ")
css = ("@page{size:A4 landscape;margin:8mm}body{font-family:'DejaVu Sans',Arial;font-size:7.2px;color:#111}h1{font-size:13px;margin:0 0 3px;color:#0B2545}h2{font-size:9.5px;margin:6px 0 2px;color:#0B2545}"
       "table{border-collapse:collapse;width:100%}th{background:#0B2545;color:#fff;padding:2px 3px;text-align:left}td{border-bottom:1px solid #ddd;padding:1px 3px;vertical-align:top}.n{text-align:right;white-space:nowrap}.g{color:#555}.neg{color:#B00020}a{color:#0563C1;text-decoration:none}")
h = [f"<!DOCTYPE html><html><head><meta charset=utf-8><style>{css}</style></head><body><h1>Маржа по закупкам РФ (ООО «Гидропас», УСН) на 15.11.2026</h1><div class=g>Прибыль после УСН 15%, без НДС. Суммы в рублях. Номер — ссылка на ЕИС.</div>",
     "<h2>1. Разобраны по присланным ТЗ</h2><table><tr><th>Номер</th><th>Заказчик / предмет</th><th>Подача</th><th>НМЦ</th><th>Себест.</th><th>При НМЦ</th><th>−10%</th><th>% −10</th><th>−20%</th><th>−10% (НДС 5% внутри)</th><th>Решение</th></tr>"]
for r in R1:
    h.append(f"<tr><td><a href='{link(r[0])}'>{r[0]}</a></td><td><b>{e(r[1])}</b> — {e(r[2])}</td><td>{e(r[3])}</td><td class=n>{k(r[4])}</td><td class=n>{k(r[5])}</td><td class=n>{k(r[6])}</td><td class=n><b>{k(r[7])}</b></td><td class=n>{r[8]}</td><td class='n{' neg' if (r[9] or 0) < 0 else ''}'>{k(r[9])}</td><td class=n>{k(r[10])}</td><td>{e(r[12])}</td></tr>")
h.append(f"</table><h2>2. Из разбора выгрузок 13.11 — подача 18.11 и позже, срок ≤6 мес ({len(R2)} закупок, по прибыли при −10%)</h2><table><tr><th>Номер</th><th>Предмет</th><th>Заказчик</th><th>НМЦ</th><th>Подача</th><th>При НМЦ</th><th>−10%</th><th>% −10</th><th>−20%</th><th>Решение 13.11</th></tr>")
for r in R2[:40]:
    h.append(f"<tr><td><a href='{link(r[0])}'>{e(r[0])}</a></td><td>{e(str(r[1])[:60])}</td><td>{e(str(r[2])[:40])}</td><td class=n>{k(r[5])}</td><td>{e(str(r[6])[:11])}</td><td class=n>{k(r[7])}</td><td class=n><b>{k(r[8])}</b></td><td class=n>{r[9] if r[9] is not None else '—'}</td><td class='n{' neg' if isinstance(r[10], (int, float)) and r[10] < 0 else ''}'>{k(r[10])}</td><td>{e(str(r[11]))}</td></tr>")
tot = sum(r[8] for r in R2 if isinstance(r[8], (int, float)) and r[8] > 0)
h.append(f"</table><div class=g>Сумма положительной прибыли при −10% по листу 2: {k(tot)} ₽. Пробирки в разборе 13.11 считались по МиниМед (реестровые записи истекли 04.09.2026) — перед подачей пересчитать по Еламед/Беримед.</div></body></html>")
hp = f"{SCRATCH}/rf_marzha.html"; open(hp, "w").write("".join(h))
subprocess.run(["/opt/pw-browsers/chromium", "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer", f"--print-to-pdf={os.path.abspath('../RF_MARZHA_1511.pdf')}", "file://" + hp], check=True, capture_output=True, timeout=180)
import pypdf; print("pages", len(pypdf.PdfReader("../RF_MARZHA_1511.pdf").pages), "R2", len(R2), "tot", tot)
for r in R1: print(r[0], r[7], r[10])
