"""ПАО «ЕГП» 32616504211 — коммутаторы Eltex: PDF на 2 стр. + Excel всех предложений.
Данные: data/zakupki_0211/tz/deep/egp_switch_prices.json (агент, 03.11.2026). → ../EGP_KOMMUTATORY_32616504211.pdf / _PREDLOZHENIYA.xlsx"""
import json, html, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from pdfutil import html_to_pdf
from tidy import tidy

P = json.load(open("data/zakupki_0211/tz/deep/egp_switch_prices.json"))
B = P["basket"]; NMC = B["nmc"]; N10 = B["nmc_minus10"]
f = lambda v: f"{v:,.0f}".replace(",", " ") if isinstance(v, (int, float)) else (v if v not in (None, "") else "—")
e = lambda s: html.escape(str(s if s is not None else "—"))
pct = lambda a, b: f"{(a - b) / a * 100:.1f} %"
def net(Pr, C, days=45):
    pre = Pr - C - 0.008 * Pr - C * 0.2 * days / 365 - 40000 - 0.015 * NMC  # доставка ~40 тыс., стоимость гарантии/заморозки 5 % обеспечения ≈1,5 % НМЦ
    return pre - 0.15 * max(pre, 0)
css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:8.6px;color:#111;margin:0}h1{font-size:14px;margin:0 0 3px}
h2{font-size:11px;margin:7px 0 3px;border-bottom:2px solid #305496;padding-bottom:1px}table{border-collapse:collapse;width:100%;margin:2px 0 5px}
th,td{border:1px solid #bbb;padding:1px 3px;vertical-align:top}th{background:#305496;color:#fff;text-align:left}td.n{text-align:right;white-space:nowrap}
.ok{background:#c6efce}.bad{background:#f8cbad}.warn{background:#fff2cc}.small{font-size:7.6px;color:#444}.box{border:1px solid #305496;padding:4px 6px;background:#f3f6fb;margin:4px 0}"""
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body>",
     "<h1>ПАО «ЕГП» — поставка коммутаторов Eltex, № 32616504211</h1>",
     "<div class='small'>ООО «ГИДРОПАС» (УСН, НДС не возмещаем — цена закупки брутто = себестоимость). Цены — открытые источники на 03.11.2026; КП дистрибьюторов нет. Полный список предложений со ссылками — в Excel.</div>",
     "<h2>1. Условия</h2><table>"]
for k, v in [("НМЦ", f"{f(NMC)} ₽ с НДС 22 %; при −10 % — {f(N10)} ₽"), ("Подача", "05.11.2026 12:00 МСК (223-ФЗ, ЕИС 32616504211)"),
             ("Предмет", "204 коммутатора Eltex, эквивалент НЕ допускается (дозакупка для совместимости); гарантия ≥12 мес. + техподдержка; по запросу — образцы на экспертизу"),
             ("Поставка / оплата", "20 рабочих дней, Ессентуки, ул. Большевистская 59А / 7 рабочих дней после УПД, аванса нет"),
             ("Обеспечение / штрафы", f"5 % цены договора (≈{f(NMC*0.05)} ₽) до подписания; штраф 1 % за нарушение"),
             ("НДС", "цена твёрдая; при утрате освобождения по ст.145 НК цена не меняется — договор на 26 млн может вывести за порог УСН без НДС"),
             ("Маркировка", "коммутаторы (ТН ВЭД 8517 62) в перечне ПП № 1954 не значатся — на дату подачи не обязательна; проверить код перед поставкой")]:
    h.append(f"<tr><th style='width:16%'>{k}</th><td>{e(v)}</td></tr>")
h.append("</table><h2>2. Цены по моделям (₽ за шт., с НДС/брутто)</h2><table><tr><th>Модель</th><th>Кол-во</th><th>Цен</th><th>Мин.</th><th>Медиана</th><th>Макс.</th><th>Σ по мин.</th><th>Σ по медиане</th><th>Соответствие ТЗ</th></tr>")
for m in P["models"]:
    tz = m.get("tz_check") or {}
    note = (str({True: "да", False: "нет"}.get(tz.get("ok"), tz.get("ok", ""))) + ": " + "; ".join(tz.get("items", [])[:2])) if isinstance(tz, dict) else str(tz)
    q = int(m["qty"])
    h.append(f"<tr><td>{e(m['model'])}</td><td class=n>{q}</td><td class=n>{m.get('n_priced')}</td><td class=n>{f(m.get('min'))}</td><td class=n>{f(m.get('median'))}</td><td class=n>{f(m.get('max'))}</td><td class=n>{f((m.get('min') or 0)*q)}</td><td class=n>{f((m.get('median') or 0)*q)}</td><td class=small>{e(note)[:230]}</td></tr>")
h.append("<tr><td>Блок питания PM160-220/12 (для MES5332A, ТЗ: ≥2 БП)</td><td class=n>4</td><td class=n>1</td><td class=n>27 195</td><td class=n>27 195</td><td class=n>27 195</td><td class=n>108 780</td><td class=n>108 780</td><td class=small>БП в комплект не входят; совместимость подтвердить у дистрибьютора</td></tr>")
h.append(f"<tr><th>Итого</th><th class=n>208</th><th></th><th></th><th></th><th></th><th class=n>{f(B['cost_min'])}</th><th class=n>{f(B['cost_median'])}</th><th>НМЦ {f(NMC)}</th></tr></table>")
h.append("<h2>3. Маржа и чистая прибыль</h2><table><tr><th>Вариант закупки</th><th>Себестоимость</th><th>Маржа по НМЦ</th><th>Маржа при −10 %</th><th>Чистая по НМЦ</th><th>Чистая при −10 %</th></tr>")
real = B.get("cost_best_real") or B.get("cost_real") or 22730000
var = [("Минимум по всем продавцам (вкл. маркетплейсы)", B["cost_min"]), ("Лучшие реальные магазины (не маркетплейсы)", real), ("Медиана рынка", B["cost_median"])] + \
      [(f"Партнёрская цена: медиана −{s['discount_pct']} %", s["cost"]) for s in P["scenarios"]]
for name, C in var:
    m10 = (N10 - C) / N10; cls = "ok" if m10 >= 0.15 else "warn" if m10 > 0 else "bad"
    h.append(f"<tr class='{cls}'><td>{e(name)}</td><td class=n>{f(C)}</td><td class=n>{pct(NMC, C)}</td><td class=n>{pct(N10, C)}</td><td class=n>{f(net(NMC, C))}</td><td class=n>{f(net(N10, C))}</td></tr>")
h.append("</table><div class='small'>Чистая: цена − закупка − 0,8 % ЭТП − доставка ≈40 тыс. − обеспечение 5 % (стоимость гарантии/заморозки ≈1,5 % НМЦ) − деньги 20 % годовых × 45 дн. − УСН 15 %.</div>")
h.append("<div class='box'><b>Вывод: НЕ ИДЁМ без КП дистрибьютора.</b> По медиане розницы корзина дороже НМЦ; даже по минимальным ценам (часть — маркетплейсы и ИП с НДС 5 %) маржа при −10 % ≈7 %. Порог 15 % при −10 % достигается только при закупке ≤ ~19,9 млн, т.е. на ~29 % ниже медианы розницы — у Eltex единая партнёрская цена для всех участников, такую скидку получить маловероятно.</div>")
h.append("<h2>4. Поставщики для КП (номер закупки, кол-ва 2/28/34/37/103 + 4 БП, проектная цена, наличие)</h2><table><tr><th>Компания</th><th>Телефон</th><th>E-mail</th><th>Сайт</th><th>ИНН</th><th>Примечание</th></tr>")
for d in (P.get("suppliers_contacts") or [])[:8]:
    h.append(f"<tr><td>{e(d.get('name'))}</td><td>{e(str(d.get('phone') or '').split(';')[0])}</td><td>{e(d.get('email'))}</td><td>{e(d.get('site'))}</td><td>{e(d.get('inn'))}</td><td class=small>{e(d.get('note'))[:120]}</td></tr>")
h.append("</table><div class='small'>Дистрибьюторы Eltex: Merlion (официальный), Treolan (партнёрское соглашение 2022), OCS — в списке партнёров. Публичных процентов партнёрских скидок нет.</div>")
h.append("<h2>5. Риски</h2><ul>" + "".join(f"<li>{e(r)[:260]}</li>" for r in P["risks"][:9]) + "</ul></body></html>")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/egp_switch.html"
open(hp, "w").write("".join(h)); html_to_pdf(hp, "../EGP_KOMMUTATORY_32616504211.pdf")

wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Все предложения"
ws.append(["Модель", "Кол-во", "Продавец", "Цена за шт., ₽", "НДС", "Исполнение", "Наличие", "Дата", "Маркетплейс", "Источник", "Ссылка"])
for c in ws[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
for m in P["models"]:
    for o in m["offers"]:
        ws.append([m["model"], m["qty"], o.get("seller"), o.get("price"), o.get("vat"), o.get("version"), o.get("stock"), o.get("date"), "да" if o.get("marketplace") else "", o.get("source"), o.get("url")])
        X = ws.max_row; ws.cell(X, 4).number_format = "#,##0"
        if o.get("url"): ws.cell(X, 11).hyperlink = o["url"]
for i, w in enumerate([18, 7, 34, 13, 26, 30, 18, 11, 9, 14, 40], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "B2"; ws.auto_filter.ref = f"A1:K{ws.max_row}"
s2 = wb.create_sheet("Итог по моделям")
s2.append(["Модель", "Кол-во", "Мин.", "Медиана", "Макс.", "Σ по мин.", "Σ по медиане"])
for m in P["models"]:
    q = int(m["qty"]); s2.append([m["model"], q, m.get("min"), m.get("median"), m.get("max"), (m.get("min") or 0) * q, (m.get("median") or 0) * q])
s2.append(["PM160-220/12 (БП для MES5332A)", 4, 27195, 27195, 27195, 108780, 108780])
s2.append(["ИТОГО", 208, None, None, None, B["cost_min"], B["cost_median"]]); s2.append(["НМЦ", None, None, None, None, NMC, NMC])
for row in s2.iter_rows(min_row=2):
    for c in row[2:]: c.number_format = "#,##0"
s3 = wb.create_sheet("Поставщики")
s3.append(["Компания", "Телефон", "E-mail", "Сайт", "ИНН", "Примечание", "2 источника"])
for d in P.get("suppliers_contacts") or []: s3.append([d.get("name"), d.get("phone"), d.get("email"), d.get("site"), d.get("inn"), d.get("note"), "да" if d.get("verified") else "нет"])
tidy(wb, row_h=22); wb.save("../EGP_KOMMUTATORY_32616504211_PREDLOZHENIYA.xlsx")
print(f(B["cost_min"]), f(B["cost_median"]), list(B.keys()))
