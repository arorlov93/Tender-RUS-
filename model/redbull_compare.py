"""Red Bull 250 мл: сводка поставщиков (ЕС, вне ЕС, КП) → себестоимость банки в Москве и маржа к опту РФ. 08.11.2026
Данные: data/companies/redbull_suppliers_eu.json, redbull_suppliers_nonEU_RU.json, КП VIE VITA, скрин Parallelbroker.
Модель — как в .claude/skills/wholesale-sourcing/scripts/landed_cost.py (растаможка в РБ). → ../REDBULL_POSTAVSHCHIKI_SVODKA.xlsx + .pdf"""
import html, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from pdfutil import html_to_pdf
from tidy import tidy

EUR, BYN = 92.1924, 27.4518
UNITS = 3564 * 24                     # 33 паллеты
RF_OPT = (101, 125)                   # опт РФ, ₽/банка (Калининская База, Граф-М, СОЦТОРГЦЕНТР, Продвей, TradeBeer) — redbull_suppliers_nonEU_RU.json
def landed(eur_can, fr_eur, sugar=True):
    p = eur_can * EUR; fr1 = fr_eur * EUR / UNITS; fr2 = 100000 / UNITS; cv = p + fr1
    duty = max(cv * .08, .03 * .25 * EUR); by_exc = .5 * .25 * BYN; fees = 63143 / UNITS; imp = (cv + duty) * .03
    rf_exc = 11 * .25 if sugar else 0
    net = p + p * .04 + fr1 + fr2 + duty + by_exc + fees + imp + 1.2 + .73 + .5 + rf_exc
    by_vat = (cv + duty + by_exc) * .2; base = cv + duty + by_exc + fees + imp + 1.2 + .73
    return net + (base + rf_exc) * .22, net + by_vat + (base + by_vat + rf_exc) * .22
# (поставщик, страна, тип, €/банка, фрахт до ЕАЭС €/фура, риск, условия, источник)
S = [
 ("Объявление flagma.de (без названия)", "DE", "B2B-объявление", .23, 3000, "высокий", "CIF", "flagma.de"),
 ("«Huster GMBH & c, GmbH» (клон Huster Getränke?)", "DE", "B2B-объявление", .42, 3000, "высокий", "DAP", "flagma"),
 ("Woove GmbH", "DE", "трейдер", .48, 3000, "высокий", "EXW", "поисковая выдача"),
 ("VIE VITA Sp. z o.o. — Sugar Free (КП № 192)", "PL", "трейдер (рег. 2001, KRS 0000004505)", .504, 2600, "средний", "40 % аванс, 60 % по документам", "КП 28.10.2026"),
 ("Ecopark Export BV", "NL", "B2B-объявление", .55, 3000, "высокий", "MOQ 33 паллеты", "flagma"),
 ("AMANOS BRAND SRL", "RO", "B2B-объявление", .55, 3200, "высокий", "EXW", "flagma, 2024–2025"),
 ("VIE VITA Sp. z o.o. — Original (КП № 192)", "PL", "трейдер (рег. 2001, KRS 0000004505)", .611, 2600, "средний", "40 % аванс, 60 % по документам", "КП 28.10.2026"),
 ("SERCO ApS", "DK", "B2B-объявление", .62, 3200, "высокий", "EXW, предоплата", "flagma"),
 ("AKLIFE IMPORT EXPORT", "TR", "B2B-объявление", .62, 5000, "высокий", "MOQ 3 564 банки", "Qoovee"),
 ("Foodica «Sp. z o.o., ApS»", "DK/PL", "B2B-объявление", .66, 3200, "высокий", "DAP, предоплата", "flagma"),
 ("Parallelbroker — лучшее из 18 (лоток 24×250)", "NL (Horst)", "площадка с эскроу", .79, 3000, "низкий", "EXW, MOQ ~29 паллет, оплата у Parallel B.V. до проверки", "parallelbroker.com"),
 ("Parallelbroker — 4-pack Classic (скрин)", "NL (Horst)", "площадка с эскроу", .91, 3000, "низкий", "EXW, MOQ 33 паллеты, 12 предложений", "скрин пользователя"),
 ("Bizim Toptan (cash & carry, Borsa Istanbul)", "TR", "cash & carry", .98, 5000, "низкий", "розничный опт", "bizimtoptan.com.tr"),
 ("Costco UK (паллета 128 кор.)", "UK", "cash & carry", 1.16, 4000, "низкий", "с доставкой по UK", "costco.co.uk"),
 ("TriStar Middle East (Emirates Bulk)", "AE", "трейдер", 1.67, 6000, "низкий", "—", "emiratesbulk"),
]
R = []
for n, c, t, pr, fr, risk, cond, src in S:
    sugar = "Sugar Free" not in n
    one, two = landed(pr, fr, sugar)
    R.append(dict(n=n, c=c, t=t, pr=pr, risk=risk, cond=cond, src=src, one=one, two=two, m1=(RF_OPT[0] - one, RF_OPT[1] - one), m2=(RF_OPT[0] - two, RF_OPT[1] - two), sugar=sugar))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Поставщики Red Bull"
ws.append(["Поставщик", "Страна", "Тип", "€ за банку", "Риск", "Условия", "Себестоимость в Москве, 1 НДС, ₽", "Себестоимость, 2 НДС, ₽", "Маржа к опту РФ 101–125 ₽ (1 НДС)", "Маржа (2 НДС)", "Источник"])
for r in R:
    ws.append([r["n"], r["c"], r["t"], r["pr"], r["risk"], r["cond"], round(r["one"], 1), round(r["two"], 1), f"{r['m1'][0]:+.0f}…{r['m1'][1]:+.0f}", f"{r['m2'][0]:+.0f}…{r['m2'][1]:+.0f}", r["src"]])
    ws.cell(ws.max_row, 5).fill = PatternFill("solid", fgColor={"низкий": "c6efce", "средний": "fff2cc"}.get(r["risk"], "f8cbad"))
s2 = wb.create_sheet("Опт РФ и РБ"); s2.append(["Продавец", "Город", "₽ за банку", "Происхождение"])
for row in [("Калининская База", "СПб", "101–106", "Австрия/Сербия"), ("Граф-М ОПТ (ОКОЛО)", "Москва", "114,80", "Австрия"), ("СОЦТОРГЦЕНТР (ОКОЛО)", "Москва", "118,71", "Австрия"),
            ("Продвей (ОКОЛО)", "Москва", "122,67", "Австрия"), ("TradeBeer", "—", "124,80", "—"), ("VendLiga", "—", "130", "—"), ("ПродПит — Sugar Free", "—", "134,79", "Австрия"),
            ("МДК ПРОДУКТ (Alligator) — минимум", "—", "91", "—"), ("Розница Беларусь", "Минск", "125–164 (4,54–5,99 BYN)", "в т.ч. Грузия")]:
    s2.append(list(row))
for s in wb:
    for c in s[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
for i, w in enumerate([44, 10, 24, 10, 10, 34, 16, 14, 18, 14, 22], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
tidy(wb, row_h=30); wb.save("../REDBULL_POSTAVSHCHIKI_SVODKA.xlsx")

e = html.escape
css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:8px;color:#111;margin:0}h1{font-size:12.5px;margin:0 0 4px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #bbb;padding:2px 3px;vertical-align:top}
th{background:#305496;color:#fff;text-align:left}td.n{text-align:right;white-space:nowrap}.ok{background:#c6efce}.warn{background:#fff2cc}.bad{background:#f8cbad}.box{border:1px solid #305496;padding:4px 6px;background:#f3f6fb;margin:3px 0}.s{font-size:7px;color:#444}"""
cl = lambda risk: {"низкий": "ok", "средний": "warn"}.get(risk, "bad")
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body><h1>Red Bull 250 мл: кто дешевле VIE VITA и что остаётся в марже (08.11.2026)</h1>",
     "<div class='box'><b>Вывод:</b> надёжных поставщиков дешевле VIE VITA (€0,61 Original / €0,50 Sugar Free) не найдено. Всё, что дешевле, — анонимные объявления (flagma, freshdi, Qoovee) высокого риска. "
     "Самая низкая проверенная цена — Parallelbroker €0,79/банка (эскроу, инспекция) — при опте в РФ 101–125 ₽/банка даёт убыток. Работать имеет смысл только по ценам ≤€0,60–0,65 и с одним НДС (НДС РБ возвращён); Sugar Free маржинальнее (нет акциза РФ). "
     "Себестоимость — 33 паллеты, растаможка в РБ, курс ЦБ €1 = 92,19 ₽; фрахт и услуги — оценка.</div>",
     "<table><tr><th>Поставщик</th><th>Страна</th><th>€/банка</th><th>Риск</th><th>Условия</th><th>Москва, 1 НДС</th><th>Москва, 2 НДС</th><th>Маржа к опту 101–125 ₽ (1 НДС)</th><th>Источник</th></tr>"]
for r in R:
    h.append(f"<tr><td>{e(r['n'])}<br><span class=s>{e(r['t'])}</span></td><td>{e(r['c'])}</td><td class=n><b>{r['pr']:.2f}</b></td><td class={cl(r['risk'])}>{e(r['risk'])}</td><td class=s>{e(r['cond'])}</td>"
             f"<td class=n>{r['one']:.0f} ₽</td><td class=n>{r['two']:.0f} ₽</td><td class='n {'ok' if r['m1'][0] > 15 else 'warn' if r['m1'][1] > 0 else 'bad'}'>{r['m1'][0]:+.0f}…{r['m1'][1]:+.0f} ₽</td><td class=s>{e(r['src'])}</td></tr>")
h.append("</table><p class=s>Опт РФ (₽/банка, 08.11.2026): Калининская База 101–106; Граф-М 114,80; СОЦТОРГЦЕНТР 118,71; Продвей 122,67; TradeBeer 124,80; ПродПит Sugar Free 134,79. Источники — data/companies/redbull_suppliers_*.json.</p></body></html>")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/rbs.html"
open(hp, "w").write("".join(h)); html_to_pdf(hp, "../REDBULL_POSTAVSHCHIKI_SVODKA.pdf")
for r in R: print(f"{r['pr']:.2f} {r['risk']:8s} {r['one']:6.1f} {r['two']:6.1f} {r['m1'][0]:+5.0f}..{r['m1'][1]:+4.0f} {r['n'][:45]}")
