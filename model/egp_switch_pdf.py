"""ПАО «ЕГП» 32616504211 — коммутаторы Eltex: досконный расчёт → ../EGP_KOMMUTATORY_32616504211.pdf
Данные: data/zakupki_0211/tz/deep/egp_switch_prices.json (цены по продавцам), egp_switch_tz.json (ТЗ)."""
import json, html, statistics
from pdfutil import html_to_pdf

D = "data/zakupki_0211/tz/deep"
P = json.load(open(f"{D}/egp_switch_prices.json"))
NMC = 26080188.33
f = lambda v: f"{v:,.0f}".replace(",", " ") if isinstance(v, (int, float)) else (v or "—")
pct = lambda v: f"{v*100:.1f} %"
e = lambda s: html.escape(str(s if s is not None else "—"))

def prof(Pr, C, days=50):
    pre = Pr - C - 0.008 * Pr - C * 0.2 * days / 365
    return pre - 0.15 * max(pre, 0)

rows = []; tot_min = tot_med = 0
for m in P["models"]:
    q = int(str(m["qty"]).split()[0]) if not isinstance(m["qty"], (int, float)) else int(m["qty"])
    mn, md, mx = m.get("min"), m.get("median"), m.get("max")
    tot_min += (mn or 0) * q; tot_med += (md or 0) * q
    rows.append((m, q, mn, md, mx))

def scen(C):
    return [pct((NMC - C) / NMC), pct((NMC * 0.9 - C) / (NMC * 0.9)), f(prof(NMC, C)), f(prof(NMC * 0.9, C))]

css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:10.5px;color:#111;margin:22px}
h1{font-size:17px;margin:0 0 4px}h2{font-size:13px;margin:16px 0 6px;border-bottom:2px solid #305496;padding-bottom:2px}
table{border-collapse:collapse;width:100%;margin:4px 0 8px}th,td{border:1px solid #bbb;padding:3px 5px;vertical-align:top}
th{background:#305496;color:#fff;text-align:left}td.n{text-align:right;white-space:nowrap}.ok{background:#c6efce}.bad{background:#f8cbad}.warn{background:#fff2cc}
.small{font-size:9px;color:#444}.box{border:1px solid #305496;padding:6px 8px;background:#f3f6fb;margin:6px 0}"""
h = [f"<html><head><meta charset='utf-8'><style>{css}</style></head><body>"]
h.append("<h1>ПАО «ЕГП» — поставка сетевого оборудования Eltex, № 32616504211</h1>")
h.append("<div class='small'>ООО «ГИДРОПАС» (УСН). Расчёт на 04.11.2026. Цены — открытые источники, с НДС (где указано). Источники и ссылки — в разделе 6.</div>")
h.append("<h2>1. Условия закупки</h2><table>")
for k, v in [("Заказчик", "ПАО «Единый гарантирующий поставщик», Ессентуки, ул. Большевистская, 59А"), ("Способ / площадка", "223-ФЗ, конкурентная закупка (извещение 32616504211)"),
             ("НМЦ", f"{f(NMC)} ₽ с НДС 22 % ({f(NMC/1.22)} ₽ без НДС)"), ("Подача заявок", "до 05.11.2026 12:00 МСК"),
             ("Предмет", "204 коммутатора Eltex, эквивалент НЕ допускается (дозакупка для совместимости)"), ("Срок поставки", "20 рабочих дней с даты договора, Ессентуки"),
             ("Оплата", "7 рабочих дней после ТОРГ-12/УПД, аванса нет"), ("Обеспечение исполнения", f"5 % цены договора (≈{f(NMC*0.05)} ₽) — гарантия или платёж до подписания"),
             ("Ответственность", "штраф 1 % цены договора за каждое нарушение (кроме просрочки); пени за просрочку"), ("Гарантия", "не менее 12 мес. + техподдержка без ограничения обращений, гарантийные карты по форме"),
             ("НДС", "цена твёрдая; при утрате освобождения по ст.145 НК цена не меняется"), ("Маркировка", e(P.get("marking")))]:
    h.append(f"<tr><th style='width:22%'>{k}</th><td>{v}</td></tr>")
h.append("</table><h2>2. Позиции и цены рынка (с НДС, за 1 шт.)</h2><table><tr><th>Модель</th><th>Кол-во</th><th>Мин.</th><th>Медиана</th><th>Макс.</th><th>Сумма по медиане</th><th>Соответствие ТЗ / замечания</th></tr>")
for m, q, mn, md, mx in rows:
    h.append(f"<tr><td>{e(m['model'])}</td><td class=n>{q}</td><td class=n>{f(mn)}</td><td class=n>{f(md)}</td><td class=n>{f(mx)}</td><td class=n>{f((md or 0)*q)}</td><td>{e(m.get('tz_check'))}</td></tr>")
h.append(f"<tr><th>Итого</th><th class=n>{sum(r[1] for r in rows)}</th><th class=n colspan=2>мин.: {f(tot_min)}</th><th></th><th class=n>{f(tot_med)}</th><th>НМЦ: {f(NMC)}</th></tr></table>")
h.append("<h2>3. Маржа и прибыль</h2><table><tr><th>Сценарий закупки</th><th>Себестоимость, ₽</th><th>Маржа по НМЦ</th><th>Маржа при −10 %</th><th>Чистая по НМЦ, ₽</th><th>Чистая при −10 %, ₽</th></tr>")
for name, C in [("Минимальные найденные цены", tot_min), ("Медиана рынка", tot_med)] + [(f"Партнёрская цена: медиана −{d} %", tot_med * (1 - d / 100)) for d in (10, 15, 20, 25)]:
    s = scen(C); cls = "ok" if (NMC*0.9 - C)/(NMC*0.9) >= 0.15 else "warn" if (NMC*0.9 - C) > 0 else "bad"
    h.append(f"<tr class='{cls}'><td>{name}</td><td class=n>{f(C)}</td><td class=n>{s[0]}</td><td class=n>{s[1]}</td><td class=n>{s[2]}</td><td class=n>{s[3]}</td></tr>")
h.append("</table><div class='small'>Чистая прибыль: цена − закупка − 0,8 % ЭТП − стоимость денег 20 % годовых × 50 дн. (поставка 20 раб. дн. + оплата 7 раб. дн.) − УСН 15 %. Доставка в Ессентуки (≈204 шт., ~1,2 т) — в пределах 30–50 тыс. ₽, в сценарии не выделена. Обеспечение 5 % — замороженные деньги/стоимость гарантии ≈1–2 % от суммы.</div>")
if P.get("scenarios"): h.append("<div class='box'><b>Сценарии агента:</b> " + e(json.dumps(P["scenarios"], ensure_ascii=False))[:1500] + "</div>")
h.append("<h2>4. Поставщики и дистрибьюторы</h2><table><tr><th>Компания</th><th>Роль</th><th>Телефон</th><th>E-mail</th><th>Сайт</th><th>ИНН</th><th>2 источника</th></tr>")
for d in P.get("distributors", []):
    h.append(f"<tr><td>{e(d.get('name'))}</td><td>{e(d.get('role') or d.get('note'))}</td><td>{e(d.get('phone'))}</td><td>{e(d.get('email'))}</td><td>{e(d.get('site'))}</td><td>{e(d.get('inn'))}</td><td>{'да' if d.get('verified') else 'нет'}</td></tr>")
h.append("</table><h2>5. Риски</h2><ul>" + "".join(f"<li>{e(r)}</li>" for r in P.get("risks", [])) + "</ul>")
h.append(f"<div class='box'><b>Вывод:</b> {e(P.get('verdict') or '')}</div>")
h.append("<h2>6. Все найденные предложения</h2><table><tr><th>Модель</th><th>Продавец</th><th>Цена</th><th>НДС</th><th>Исполнение</th><th>Наличие</th><th>Ссылка</th></tr>")
for m, q, *_ in rows:
    for o in m.get("offers", []):
        h.append(f"<tr><td>{e(m['model'])}</td><td>{e(o.get('seller'))}</td><td class=n>{f(o.get('price'))}</td><td>{e(o.get('vat'))}</td><td>{e(o.get('version'))}</td><td>{e(o.get('stock'))}</td><td class=small>{e(o.get('url'))}</td></tr>")
h.append("</table></body></html>")
open("/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/egp_switch.html", "w").write("".join(h))
html_to_pdf("/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/egp_switch.html", "../EGP_KOMMUTATORY_32616504211.pdf")
print(f(tot_min), f(tot_med))
