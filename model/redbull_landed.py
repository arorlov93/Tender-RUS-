"""Red Bull 250 мл: расчёт себестоимости в Москве (Литва → Беларусь → Москва), 08.11.2026 → ../REDBULL_RASCHET_0811.xlsx + .pdf
Все допущения — в таблице DOP. Курс ЦБ на 07.11.2026: EUR 92,1924 ₽."""
import html, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from pdfutil import html_to_pdf
from tidy import tidy

EUR = 92.1924
PAL, CASES, CANS, L = 26, 108, 24, 0.25           # 26 паллет в фуре (лимит ≈20 т: паллета ≈750 кг)
N = PAL * CASES * CANS                              # банок в фуре
DOP = [  # (статья, значение, ед., комментарий/источник)
 ("Паллет в фуре", PAL, "шт", "паллета 108 коробок × 24 банки ≈ 750 кг; фура ≈20 т"),
 ("Банок в фуре", N, "шт", f"{N/1000:.1f} тыс. банок = {N*L:,.0f} л".replace(",", " ")),
 ("Курс EUR", EUR, "₽", "ЦБ РФ на 07.11.2026"),
 ("Платёжный агент / банковские расходы", 0.04, "доля", "оплата через агента в третьей стране, 3–5 % — оценка"),
 ("Фрахт Литва → Минск", 2200, "€/фура", "перевозчик ЕС + очереди на границе — оценка"),
 ("Фрахт Минск → Москва", 80000, "₽/фура", "оценка"),
 ("Пошлина ЕАЭС, код 2202 99 180 0", 0.08, "доля", "8 %, но не менее 0,03 €/л (tws.by — сверить с ЕТТ)"),
 ("Таможенный сбор + брокер + СВХ в Беларуси", 55000, "₽/фура", "оценка"),
 ("Вознаграждение белорусского импортёра", 0.03, "доля", "оценка, от стоимости товара"),
 ("Акциз РФ на сахаросодержащие напитки", 11, "₽/л", "11 ₽/л с 2026 г. (dairynews; сверить со ст. 193 НК) — только Original; Sugarfree/Zero — 0"),
 ("Акциз Беларуси на энергетики", 0.5, "BYN/л", "с 2026 г.; при вывозе в РФ, вероятно, не взимается — проверить"),
 ("НДС при ввозе в РФ из ЕАЭС", 0.22, "доля", "для ОСНО — к вычету; для УСН — в себестоимость"),
 ("Код «Честный знак»", 0.73, "₽/банка", "0,60 ₽ + НДС"),
 ("Стикеровка на русском", 1.2, "₽/банка", "оценка; этикетки EN/DE"),
 ("Декларация соответствия ЕАЭС", 40000, "₽/партия", "оценка"),
 ("Склад/обработка в Москве", 0.5, "₽/банка", "оценка"),
]
def calc(price_eur, sugar=True):
    goods = price_eur * EUR
    agent = goods * .04
    fr1 = 2200 * EUR / N; fr2 = 80000 / N
    cv = goods + fr1                                   # таможенная стоимость до границы ЕАЭС
    duty = max(cv * .08, .03 * L * EUR)
    fees = 55000 / N; imp = goods * .03
    exc = 11 * L if sugar else 0
    cz, stick, cert, wh = .73, 1.2, 40000 / N, .5
    net = goods + agent + fr1 + fr2 + duty + fees + imp + exc + cz + stick + cert + wh
    vat = (cv + duty + exc) * .22
    return dict(goods=goods, agent=agent, fr=fr1 + fr2, duty=duty, fees=fees + imp, exc=exc, mark=cz + stick + cert + wh, osno=net, vat=vat, usn=net + vat)
SC = [("€0,30 — Original (с сахаром)", .30, True), ("€0,30 — Sugarfree/Zero", .30, False), ("€0,40 — Original", .40, True), ("€0,40 — Sugarfree/Zero", .40, False)]
R = [(n, calc(p, s)) for n, p, s in SC]
ROWS = [("Товар (цена поставщика × курс)", "goods"), ("Платёжный агент 4 %", "agent"), ("Фрахт Литва → Минск → Москва", "fr"), ("Пошлина ЕАЭС 8 %", "duty"),
        ("Сбор, брокер, СВХ, импортёр РБ", "fees"), ("Акциз РФ (сахар)", "exc"), ("Маркировка, стикер, сертификат, склад", "mark"),
        ("ИТОГО без НДС (ОСНО — НДС к вычету)", "osno"), ("НДС при ввозе 22 %", "vat"), ("ИТОГО для УСН (НДС в себестоимости)", "usn")]

wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Себестоимость банки"
ws.append(["₽ за банку 250 мл"] + [n for n, _ in R])
for t, k in ROWS:
    ws.append([t] + [round(r[k], 2) for _, r in R])
    if k in ("osno", "usn"):
        for c in ws[ws.max_row]: c.font = Font(bold=True)
ws.append([]); ws.append(["За фуру 26 паллет (67 392 банки), млн ₽ — УСН"] + [round(r["usn"] * N / 1e6, 2) for _, r in R])
ws.append(["Оборотка на фуру: товар + доставка + пошлины, млн ₽"] + [round((r["goods"] + r["agent"] + r["fr"] + r["duty"] + r["fees"]) * N / 1e6, 2) for _, r in R])
s2 = wb.create_sheet("Допущения"); s2.append(["Статья", "Значение", "Ед.", "Источник / комментарий"])
for row in DOP: s2.append(list(row))
for s in wb:
    for c in s[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
ws.column_dimensions["A"].width = 48
for col in "BCDE": ws.column_dimensions[col].width = 20
for col, w in zip("ABCD", (42, 14, 12, 70)): s2.column_dimensions[col].width = w
tidy(wb, row_h=20); wb.save("../REDBULL_RASCHET_0811.xlsx")

e = html.escape
f2 = lambda v: f"{v:,.2f}".replace(",", " ")
css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:9px;color:#111;margin:0}h1{font-size:13px;margin:0 0 4px}h2{font-size:10.5px;margin:6px 0 3px;border-bottom:2px solid #305496}table{border-collapse:collapse;width:100%}
th,td{border:1px solid #bbb;padding:2px 4px}th{background:#305496;color:#fff;text-align:left}td.n{text-align:right}.b{font-weight:bold;background:#f3f6fb}.box{border:1px solid #305496;padding:5px 7px;background:#f3f6fb;margin:4px 0}"""
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body><h1>Red Bull 250 мл: сколько стоит банка в Москве (Литва → Беларусь → Москва)</h1>",
     f"<div class='box'>Фура 26 паллет = {N:,} банок. Курс ЦБ 07.11.2026: €1 = {EUR} ₽. Растаможка в Беларуси через белорусского импортёра, ввоз в РФ из ЕАЭС. Ставки пошлины и акциза — из открытых источников, сверить с брокером; фрахт, агент, стикеровка — оценки.</div>".replace(",", " "),
     "<table><tr><th>₽ за банку</th>" + "".join(f"<th>{e(n)}</th>" for n, _ in R) + "</tr>"]
for t, k in ROWS:
    cls = " class=b" if k in ("osno", "usn") else ""
    h.append(f"<tr{cls}><td>{e(t)}</td>" + "".join(f"<td class=n>{f2(r[k])}</td>" for _, r in R) + "</tr>")
h.append("<tr class=b><td>Фура, млн ₽ (УСН)</td>" + "".join(f"<td class=n>{r['usn']*N/1e6:.2f}</td>" for _, r in R) + "</tr></table>")
h.append("<h2>Допущения</h2><table><tr><th>Статья</th><th>Значение</th><th>Ед.</th><th>Источник</th></tr>" + "".join(f"<tr><td>{e(a)}</td><td class=n>{e(str(b))}</td><td>{e(c)}</td><td>{e(d)}</td></tr>" for a, b, c, d in DOP) + "</table></body></html>")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/rb.html"
open(hp, "w").write("".join(h)); html_to_pdf(hp, "../REDBULL_RASCHET_0811.pdf")
for n, r in R: print(n, {k: round(v, 2) for k, v in r.items()})
