"""Red Bull по КП VIE VITA SP Z O.O. (Польша) № 192 от 28.10.2026 для ISP Group LLC: расчёт банки в Москве (Польша → Беларусь, растаможка в РБ → Москва)
→ ../REDBULL_VIEVITA_RASCHET.xlsx + .pdf. Курс ЦБ РФ 07.11.2026: EUR 92,1924; BYN 27,4518."""
import html, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from pdfutil import html_to_pdf
from tidy import tidy

EUR, BYN = 92.1924, 27.4518
ITEMS = [("Red Bull 250 мл Original", 1782, 14.66, True), ("Red Bull 250 мл Sugar Free", 1782, 12.10, False)]
CASES = sum(i[1] for i in ITEMS); CANS = CASES * 24; LIT = CANS * .25; PAL = CASES / 108
FR_PL_BY = 2600      # € Кельце → Минск, одна тяжёлая фура (≈24 т), перевозчик ЕС — оценка
FR_BY_MSK = 100000   # ₽ Минск → Москва — оценка
BY_FIX = (100 + 400 + 300 + 1500) * BYN   # сбор + брокер + СВХ + декларация ЕАЭС, BYN — оценка

def per_can(price_case, sugar):
    p = price_case / 24 * EUR                       # цена банки, ₽
    bank = p * .04
    fr1 = FR_PL_BY * EUR / CANS; fr2 = FR_BY_MSK / CANS
    cv = p + fr1                                    # таможенная стоимость на границе ЕАЭС
    duty = max(cv * .08, .03 * .25 * EUR)
    by_exc = .5 * .25 * BYN                          # акциз РБ на энергетики
    fees = BY_FIX / CANS; imp = (cv + duty) * .03
    stick, cz, wh = 1.2, .73, .5
    rf_exc = 11 * .25 if sugar else 0
    net = p + bank + fr1 + fr2 + duty + by_exc + fees + imp + stick + cz + wh + rf_exc
    by_vat = (cv + duty + by_exc) * .20
    base_rf = cv + duty + by_exc + fees + imp + stick + cz
    rf_vat1 = (base_rf + rf_exc) * .22                     # один НДС (НДС РБ вернули, продажа в РФ по 0 %)
    rf_vat2 = (base_rf + by_vat + rf_exc) * .22            # два НДС (НДС РБ остался в цене)
    return dict(price=p, bank=bank, freight=fr1 + fr2, duty=duty, by_exc=by_exc, fees=fees + imp, mark=stick + cz + wh, rf_exc=rf_exc, net=net,
                by_vat=by_vat, rf_vat1=rf_vat1, rf_vat2=rf_vat2, one=net + rf_vat1, two=net + by_vat + rf_vat2)
R = [(n, q, pc, per_can(pc, s)) for n, q, pc, s in ITEMS]
ROWS = [("Цена поставщика за банку (€/24 × курс)", "price"), ("Банк / агент 4 %", "bank"), ("Доставка Польша → Минск → Москва", "freight"), ("Пошлина ЕАЭС 8 %", "duty"),
        ("Акциз РБ 0,5 BYN/л", "by_exc"), ("Брокер, сбор, СВХ, декларация, импортёр РБ", "fees"), ("Стикер, «Честный знак», склад", "mark"), ("Акциз РФ 11 ₽/л (только Original)", "rf_exc"),
        ("ИТОГО без НДС", "net"), ("НДС РФ 22 % (один НДС)", "rf_vat1"), ("ИТОГО с одним НДС (НДС РБ вернули)", "one"), ("НДС РБ 20 % (не вернули)", "by_vat"),
        ("НДС РФ 22 % (с учётом НДС РБ)", "rf_vat2"), ("ИТОГО с двумя НДС", "two")]

wb = openpyxl.Workbook(); ws = wb.active; ws.title = "₽ за банку"
ws.append(["₽ за банку 250 мл"] + [f"{n} (€{pc}/кор.)" for n, q, pc, _ in R] + ["Весь заказ, ₽"])
for t, k in ROWS:
    tot = sum(r[k] * q * 24 for n, q, pc, r in R)
    ws.append([t] + [round(r[k], 2) for *_, r in R] + [round(tot)])
    if k in ("net", "one", "two"):
        for c in ws[ws.max_row]: c.font = Font(bold=True)
ws.append([]); ws.append(["Заказ: коробок / банок / паллет / литров", f"{CASES} / {CANS} / {PAL:.1f} / {LIT:.0f}"])
ws.append(["Сумма КП, €", round(sum(q * pc for n, q, pc, _ in R), 2)])
ws.append(["Оплата по КП", "40 % TT после инвойса, 60 % против отгрузочных документов и копии B/L"])
for c in ws[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
ws.column_dimensions["A"].width = 50
for col in "BCD": ws.column_dimensions[col].width = 26
tidy(wb, row_h=20); wb.save("../REDBULL_VIEVITA_RASCHET.xlsx")

e = html.escape; f2 = lambda v: f"{v:,.2f}".replace(",", " "); f0 = lambda v: f"{v:,.0f}".replace(",", " ")
css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:9px;color:#111;margin:0}h1{font-size:13px;margin:0 0 4px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #bbb;padding:2px 4px}
th{background:#305496;color:#fff;text-align:left}td.n{text-align:right}.b{font-weight:bold;background:#f3f6fb}.box{border:1px solid #305496;padding:5px 7px;background:#f3f6fb;margin:4px 0}"""
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body><h1>Red Bull по КП VIE VITA (Польша) № 192: сколько стоит банка в Москве</h1>",
     f"<div class='box'>Заказ: {CASES} коробок = {CANS:,} банок ≈ {PAL:.0f} паллет (≈24 т, одна тяжёлая фура). Сумма КП €{sum(q*pc for n,q,pc,_ in R):,.2f}. Маршрут: Кельце → Минск (растаможка в РБ через импортёра) → Москва. Курс ЦБ 07.11.2026 €1 = {EUR} ₽. Доставка, брокер, агент — оценка.</div>".replace(",", " "),
     "<table><tr><th>₽ за банку</th>" + "".join(f"<th>{e(n)}<br>€{pc}/коробка</th>" for n, q, pc, _ in R) + "<th>Весь заказ, ₽</th></tr>"]
for t, k in ROWS:
    tot = sum(r[k] * q * 24 for n, q, pc, r in R)
    h.append(f"<tr{' class=b' if k in ('net','one','two') else ''}><td>{e(t)}</td>" + "".join(f"<td class=n>{f2(r[k])}</td>" for *_, r in R) + f"<td class=n>{f0(tot)}</td></tr>")
h.append("</table></body></html>")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/vv.html"
open(hp, "w").write("".join(h)); html_to_pdf(hp, "../REDBULL_VIEVITA_RASCHET.pdf")
for n, q, pc, r in R: print(n, pc, {k: round(r[k], 2) for k in ("price", "net", "one", "two")})
print("заказ ₽: net", round(sum(r['net']*q*24 for n,q,pc,r in R)), "one", round(sum(r['one']*q*24 for n,q,pc,r in R)), "two", round(sum(r['two']*q*24 for n,q,pc,r in R)), CANS, PAL)
