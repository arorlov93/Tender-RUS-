"""Red Bull 250 мл: полный расчёт при растаможке в Беларуси (импортёр РБ → поставка в РФ) против транзита и растаможки в РФ. 08.11.2026
→ ../REDBULL_RB_RASTAMOZHKA_0811.xlsx + .pdf. Курсы ЦБ РФ 07.11.2026: EUR 92,1924; BYN 27,4518."""
import html, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from pdfutil import html_to_pdf
from tidy import tidy

EUR, BYN = 92.1924, 27.4518
PAL = 26; N = PAL * 108 * 24; L = 0.25; LIT = N * L                     # фура: 67 392 банки, 16 848 л
P = dict(agent=.04, fr_lt=2200, fr_by_msk=80000, duty=.08, duty_min=.03, by_excise=.5, by_vat=.20, by_fee=100, broker=400, svh=300, cert=1500,
         by_imp=.03, rf_exc=11, rf_vat=.22, cz=.73, stick=1.2, wh=.5, fin_rate=.18, vat_months=2.5, rf_broker=60000, rf_fee=8530)

def run(price, mode, sugar=True):
    """mode: 'BY_refund' — растаможка в РБ, акциз РБ возвращается/не взимается при вывозе в РФ; 'BY_norefund' — акциз РБ остаётся в цене; 'RF' — транзит через РБ, растаможка в РФ."""
    g = price * EUR; ag = g * P["agent"]; fr1 = P["fr_lt"] * EUR / N; fr2 = P["fr_by_msk"] / N
    cv = g + fr1; duty = max(cv * P["duty"], P["duty_min"] * L * EUR)
    rf_exc = P["rf_exc"] * L if sugar else 0
    o = dict(goods=g, agent=ag, freight=fr1 + fr2, duty=duty)
    if mode.startswith("BY"):
        by_exc = P["by_excise"] * L * BYN
        by_vat = (cv + duty + by_exc) * P["by_vat"]
        o["by_fees"] = (P["by_fee"] + P["broker"] + P["svh"] + P["cert"]) * BYN / N
        o["by_importer"] = (cv + duty) * P["by_imp"]
        o["by_excise"] = 0 if mode == "BY_refund" else by_exc
        o["by_vat_freeze"] = by_vat * P["fin_rate"] * P["vat_months"] / 12        # НДС РБ при ввозе возмещается после экспорта в РФ (0 %) — стоимость денег
        o["rf_fees"] = 0
        base_rf = cv + duty + o["by_fees"] + o["by_importer"] + o["by_excise"] + fr2
        o["_by_cash"] = by_vat + by_exc                                           # временно замороженные деньги в РБ на банку
    else:
        o["by_fees"] = 30000 / N                                                  # транзит/перегруз на границе РБ
        o["by_importer"] = 0; o["by_excise"] = 0; o["by_vat_freeze"] = 0
        o["rf_fees"] = (P["rf_broker"] + P["rf_fee"]) / N
        base_rf = cv + duty; o["_by_cash"] = 0
    o["rf_excise"] = rf_exc
    o["mark"] = P["cz"] + P["stick"] + P["wh"] + (0 if mode.startswith("BY") else P["cert"] * BYN / N)
    o["net"] = sum(v for k, v in o.items() if not k.startswith("_"))
    o["rf_vat"] = (base_rf + rf_exc) * P["rf_vat"]
    o["usn"] = o["net"] + o["rf_vat"]
    return o
MODES = [("BY_refund", "РБ, акциз РБ не остаётся"), ("BY_norefund", "РБ, акциз РБ остаётся"), ("RF", "Транзит, растаможка в РФ")]
SC = [(p, m, t) for p in (.30, .40) for m, t in MODES]
R = [(f"€{p:.2f} · {t}".replace(".", ","), run(p, m)) for p, m, t in SC]
ROWS = [("Товар", "goods"), ("Платёжный агент 4 %", "agent"), ("Фрахт Литва → Минск → Москва", "freight"), ("Пошлина ЕАЭС 8 %", "duty"),
        ("Таможня/брокер/СВХ/сертификат в РБ (или перегруз)", "by_fees"), ("Вознаграждение импортёра РБ 3 %", "by_importer"), ("Акциз РБ 0,5 BYN/л", "by_excise"),
        ("Стоимость заморозки НДС РБ 20 % (≈2,5 мес.)", "by_vat_freeze"), ("Брокер/сбор в РФ", "rf_fees"), ("Акциз РФ 11 ₽/л (Original)", "rf_excise"),
        ("«Честный знак», стикер, склад", "mark"), ("ИТОГО без НДС РФ (ОСНО)", "net"), ("НДС РФ 22 % при ввозе", "rf_vat"), ("ИТОГО для УСН", "usn")]

wb = openpyxl.Workbook(); ws = wb.active; ws.title = "₽ за банку"
ws.append(["₽ за банку 250 мл (Original)"] + [n for n, _ in R])
for t, k in ROWS:
    ws.append([t] + [round(r[k], 2) for _, r in R])
    if k in ("net", "usn"):
        for c in ws[ws.max_row]: c.font = Font(bold=True)
ws.append([]); ws.append(["Фура 26 паллет, млн ₽ (УСН)"] + [round(r["usn"] * N / 1e6, 2) for _, r in R])
ws.append(["Деньги временно заморожены в РБ (НДС + акциз РБ), млн ₽/фура"] + [round(r["_by_cash"] * N / 1e6, 2) for _, r in R])
s2 = wb.create_sheet("Допущения")
s2.append(["Параметр", "Значение", "Комментарий"])
for k, v, c in [("Курс EUR / BYN (ЦБ РФ 07.11.2026)", f"{EUR} / {BYN}", "cbr-xml-daily.ru"), ("Фура", f"{PAL} паллет, {N} банок, {LIT:.0f} л", "≈750 кг паллета, ≈20 т фура"),
                ("Пошлина ЕАЭС", "8 %, не менее 0,03 €/л", "код 2202 99 180 0 (tws.by) — сверить с ЕТТ"), ("Акциз РБ на энергетики", "0,5 BYN/л с 01.01.2026", "myfin.by; возврат при вывозе в РФ — уточнить у брокера РБ"),
                ("НДС РБ при ввозе", "20 %", "к вычету/возмещению у импортёра РБ после экспорта в РФ по 0 %"), ("Акциз РФ сахаросодержащие", "11 ₽/л", "только Original; Sugarfree/Zero — 0; сверить со ст. 193 НК"),
                ("НДС РФ при ввозе из ЕАЭС", "22 %", "для УСН — в себестоимость"), ("Фрахт", "€2 200 Литва → Минск; 80 000 ₽ Минск → Москва", "оценка"),
                ("Брокер, СВХ, сбор, декларация ЕАЭС в РБ", "100 + 400 + 300 + 1 500 BYN", "оценка"), ("Импортёр РБ", "3 %", "оценка"), ("Платёжный агент", "4 %", "оценка"),
                ("Стоимость денег", "18 % годовых", "заморозка НДС РБ на 2,5 мес.")]:
    s2.append([k, v, c])
for s in wb:
    for c in s[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
ws.column_dimensions["A"].width = 52
for i in range(2, 8): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 19
for col, w in zip("ABC", (40, 40, 70)): s2.column_dimensions[col].width = w
tidy(wb, row_h=30); wb.save("../REDBULL_RB_RASTAMOZHKA_0811.xlsx")

e = html.escape; f2 = lambda v: f"{v:,.2f}".replace(",", " ")
css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:8.3px;color:#111;margin:0}h1{font-size:12.5px;margin:0 0 4px}h2{font-size:10px;margin:6px 0 3px;border-bottom:2px solid #305496}table{border-collapse:collapse;width:100%}
th,td{border:1px solid #bbb;padding:2px 3px}th{background:#305496;color:#fff;text-align:left}td.n{text-align:right}.b{font-weight:bold;background:#f3f6fb}.box{border:1px solid #305496;padding:4px 6px;background:#f3f6fb;margin:3px 0}"""
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body><h1>Red Bull 250 мл: растаможка в Беларуси — полный расчёт (₽ за банку, Москва)</h1>",
     "<div class='box'><b>Схема РБ:</b> поставщик (Литва) → импортёр РБ растаможивает в Минске (пошлина 8 %, акциз РБ 0,5 BYN/л, НДС РБ 20 %) → продаёт в РФ по 0 % НДС → покупатель в РФ платит НДС 22 % и акциз РФ 11 ₽/л по заявлению о ввозе из ЕАЭС. "
     "<b>Схема РФ:</b> транзит через РБ, растаможка в Москве (пошлина 8 %, НДС 22 %, акциз РФ). Курсы ЦБ 07.11.2026. Ставки — из открытых источников, сверить с брокером РБ; фрахт и услуги — оценки.</div>",
     "<table><tr><th>₽ за банку</th>" + "".join(f"<th>{e(n)}</th>" for n, _ in R) + "</tr>"]
for t, k in ROWS:
    h.append(f"<tr{' class=b' if k in ('net','usn') else ''}><td>{e(t)}</td>" + "".join(f"<td class=n>{f2(r[k])}</td>" for _, r in R) + "</tr>")
h.append("<tr class=b><td>Фура, млн ₽ (УСН)</td>" + "".join(f"<td class=n>{r['usn']*N/1e6:.2f}</td>" for _, r in R) + "</tr>")
h.append("<tr><td>Заморожено в РБ (НДС + акциз), млн ₽/фура</td>" + "".join(f"<td class=n>{r['_by_cash']*N/1e6:.2f}</td>" for _, r in R) + "</tr></table>")
h.append("""<h2>Что важно при растаможке в РБ</h2><ul>
<li><b>Нужен импортёр-резидент РБ</b> (ООО/ИП): он подписывает контракт с поставщиком, платит в ЕС, растаможивает, получает декларацию ЕАЭС (ТР ТС 021/022) и продаёт вам. Платежи из РБ в ЕС тоже затруднены — возможно, понадобится агент в Казахстане/ОАЭ.</li>
<li><b>Акциз РБ на энергетики с 2026 г.</b> — платится при ввозе. Возвращается ли при вывозе в РФ — уточнить у брокера; если нет — +3,4 ₽/банка.</li>
<li><b>НДС РБ 20 %</b> — платится на таможне и возмещается импортёру после экспорта в РФ (0 %) — заморозка ≈0,7–0,9 млн ₽ на фуру на 2–3 месяца.</li>
<li><b>Маркировка на русском</b> (ТР ТС 022) обязательна уже для обращения в РБ — стикеровка до выпуска/на складе в РБ. «Честный знак» РФ — коды нанести до въезда в РФ; проверить требования маркировки РБ.</li>
<li><b>Ввоз в РФ из ЕАЭС:</b> заявление о ввозе + НДС 22 % и акциз РФ до 20-го числа следующего месяца; транспорт РБ → РФ без таможенного оформления.</li>
<li><b>Литва → Беларусь:</b> граница периодически закрывается; перевозчики РБ/РФ в ЕС не допускаются — везёт перевозчик ЕС до Минска или перегруз на границе.</li></ul></body></html>""")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/rb2.html"
open(hp, "w").write("".join(h)); html_to_pdf(hp, "../REDBULL_RB_RASTAMOZHKA_0811.pdf")
for n, r in R: print(n, round(r["net"], 2), round(r["usn"], 2), round(r["usn"] * N / 1e6, 2), round(r["_by_cash"] * N / 1e6, 2))
