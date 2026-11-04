"""ПАО «Россети Центр и Приволжье» — Удмуртэнерго, ЗК МСП 32616535735: оборудование для дистанц. контроля доступа на РП (28 поз.).
Данные: deep/svet_items.json (ТЗ + НМЦ ед. из скана обоснования), deep/svet_out.json (цены, агент 04.11.2026). → ../CHELYAB_SVET_32616505090.pdf + _POZICII.xlsx"""
import json, html, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from pdfutil import html_to_pdf
from tidy import tidy

D = "data/zakupki_0211/tz/deep/"
IT = {x["n"]: x for x in json.load(open(D + "svet_items.json"))}
O = json.load(open(D + "svet_out.json")); R = {x["n"]: x for x in O["items"]}
NMC = 922272.08; NMC_NV = 755960.72
DELIV = 12000; DAYS = 65
f = lambda v: f"{v:,.0f}".replace(",", " ") if isinstance(v, (int, float)) else "—"
f2 = lambda v: f"{v:,.2f}".replace(",", " ") if isinstance(v, (int, float)) else "—"
e = lambda s: html.escape(str(s if s is not None else "—"))
rows = []
for n in sorted(IT):
    x = IT[n]; r = R.get(n, {}); up = r.get("unit_price"); est = r.get("est_unit_price")
    u = up if isinstance(up, (int, float)) else est if isinstance(est, (int, float)) else None
    rows.append(dict(x, regime="ограничение" if n in (17, 18, 24, 25, 26, 27, 30) else "—", up=up, est=est, u=u, cost=(u or 0) * x["qty"], nmc_sum=x["nmc_sum"], prod=r.get("product"), reg=r.get("registry"), match=r.get("match"), note=r.get("note"), src=r.get("sources") or [], vat=r.get("vat")))
C_A = sum(r["cost"] for r in rows) + sum(r["nmc_sum"] for r in rows if not r["u"])     # без цены — по НМЦ
C_B = C_A - (rows[29]["nmc_sum"] if not rows[29]["u"] else 0) + 202514 * 2 if not rows[29]["u"] else C_A            # поз. 30 по ориентиру РУ-ТТ 202 514
C_conf = sum(r["up"] * r["qty"] for r in rows if isinstance(r["up"], (int, float)))
def net(P, C):
    pre = P - C - DELIV - 0.008 * P - C * 0.2 * DAYS / 365
    return pre - 0.15 * max(pre, 0)
css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:8.3px;color:#111;margin:0}h1{font-size:14px;margin:0 0 3px}
h2{font-size:11px;margin:6px 0 3px;border-bottom:2px solid #305496;padding-bottom:1px}table{border-collapse:collapse;width:100%;margin:2px 0 4px}
th,td{border:1px solid #bbb;padding:1px 3px;vertical-align:top}th{background:#305496;color:#fff;text-align:left}td.n{text-align:right;white-space:nowrap}
.ok{background:#c6efce}.bad{background:#f8cbad}.warn{background:#fff2cc}.small{font-size:7.2px;color:#444}.box{border:1px solid #305496;padding:4px 6px;background:#f3f6fb;margin:4px 0}"""
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body>",
     "<h1>ПАО «Россети Урал» — «Челябэнерго»: прочее светотехническое оборудование, ЕИС 32616505090 (ЗП, ID 62097)</h1>",
     "<div class='small'>ООО «ГИДРОПАС» (УСН). Извещение, документация, ТЗ, опросные листы, расчёт НМЦ. Цены — открытые источники на 04.11.2026, с НДС, без доставки; КП не запрашивались; реестр ПП 719 — ГИСП 01.11.2026. Ссылки — в Excel.</div><h2>1. Условия</h2><table>"]
for a, b in [("НМЦ", f"{f2(NMC)} ₽ с НДС ({f2(NMC_NV)} без НДС); 35 позиций, частичная подача не допускается"),
             ("Подача", "запрос предложений (для всех, не только МСП); corp.roseltorg.ru до 09.11.2026 09:00 МСК; рассмотрение 17.11; переторжка 19.11"),
             ("Оценка", "цена 70 %, аттестация ПАО «Россети» 20 %, опыт 5 %, полномочия производителей 5 %"),
             ("Поставка", "Челябинская обл., 30 календарных дней с заключения договора"),
             ("Оплата", "30 рабочих дней после приёмки"),
             ("Нацрежим", "ОГРАНИЧЕНИЕ — поз. 17, 18, 24–27, 30 (нужен номер записи ПП 719)")]:
    h.append(f"<tr><th style='width:11%'>{a}</th><td>{e(b)}</td></tr>")
h.append("</table><h2>2. Позиции: НМЦ против рынка (с НДС)</h2><table><tr><th>№</th><th>Позиция</th><th>Кол.</th><th>НМЦ ед.</th><th>Рынок ед.</th><th>НМЦ сумма</th><th>Рынок сумма</th><th>Рынок/НМЦ</th><th>Режим / реестр</th><th>Товар</th></tr>")
for r in sorted(rows, key=lambda r: -r["nmc_sum"]):
    q = r["cost"] / r["nmc_sum"] if r["u"] else None
    cls = "bad" if q is None or q > 1 else "warn" if q > 0.8 else "ok"
    lab = f2(r["u"]) + ("*" if not isinstance(r["up"], (int, float)) and r["u"] else "")
    h.append(f"<tr class='{cls}'><td>{r['n']}</td><td>{e(r['name'][:60])}</td><td class=n>{r['qty']:g}</td><td class=n>{f2(r['nmc_unit'])}</td><td class=n>{lab}</td><td class=n>{f(r['nmc_sum'])}</td><td class=n>{f(r['cost'])}</td>"
             f"<td class=n>{'—' if q is None else f'{q*100:.0f} %'}</td><td class=small>{e(r['regime'][:13])}{'; ' + e(str(r['reg'])[:22]) if r['reg'] else ''}</td><td class=small>{e(str(r['prod'])[:70])}</td></tr>")
h.append(f"<tr><th colspan=5>Итого найденных (без цены: поз. 7, 9, 20, 30, 33)</th><th class=n>{f(NMC)}</th><th class=n>{f(C_B)}</th><th class=n>{C_B/NMC*100:.0f} %</th><th colspan=2>подтверждённые цены: {f(C_conf)} ₽</th></tr></table>")
h.append("<h2>3. Маржа (сравнение без НДС: наша цена ≤ НМЦ без НДС, закупка с НДС)</h2><table><tr><th>Вариант</th><th>Себестоимость + доставка</th><th>Маржа по НМЦ</th><th>При −10 %</th><th>Чистая по НМЦ</th><th>Чистая при −10 %</th></tr>")
for nm, C in [("А: позиции без цены (7, 9, 20, 30, 33) — по НМЦ", C_A), ("Б: то же, поз. 30 БЭМП — по ориентиру младшей модели 202 514 ₽/шт", C_B)]:
    m1 = (NMC_NV * .9 - C - DELIV) / (NMC_NV * .9)
    h.append(f"<tr class='{'ok' if m1 >= .15 else 'warn' if m1 > 0 else 'bad'}'><td>{e(nm)}</td><td class=n>{f(C + DELIV)}</td><td class=n>{(NMC_NV - C - DELIV) / NMC_NV * 100:.1f} %</td><td class=n>{m1*100:.1f} %</td><td class=n>{f(net(NMC_NV, C))}</td><td class=n>{f(net(NMC_NV * .9, C))}</td></tr>")
h.append(f"</table><div class='small'>Доставка ≈{f(DELIV)} ₽ (оценка). Чистая: цена − закупка − доставка − 0,8 % ЭТП − деньги 20 % год. × {DAYS} дн. − УСН 15 %.</div>")
h.append(f"<div class='box'><b>Вывод: НЕ ИДЁМ.</b> Даже по открытым ценам 30 найденных позиций дороже их НМЦ на {(sum(r['cost'] for r in rows if r['u'])/sum(r['nmc_sum'] for r in rows if r['u'])-1)*100:.0f} % (РУ-21, РЭП36Н, Э42702, ЗН27 — все выше НМЦ), а две главные позиции (БЭМП РУ-ТТ3 ЧЭАЗ — 2×193 тыс., комплект втычных контактов «Таврида» ВЭ/TEL — 292 тыс., вместе ≈74 % НМЦ) — только у единственных производителей по КП; ориентир по БЭМП не ниже НМЦ. "
         "Наша цена сравнивается с НМЦ без НДС ({f(NMC_NV)}), а закупаем с НДС — минус 18 % сразу. 20 % оценки — аттестация Россети, у нас её нет. PM130EH PLUS (SATEC, Израиль) под ограничением; российский ЩМ120 дороже НМЦ и не подтверждён по ОЛ полностью.</div>")
h.append("<h2>4. Поставщики для КП</h2><table><tr><th>Компания</th><th>Телефон</th><th>E-mail</th><th>Позиции</th></tr>")
for s in O["suppliers"]:
    h.append(f"<tr><td>{e(s.get('name'))[:80]}</td><td>{e(s.get('phone'))[:40]}</td><td>{e(s.get('email'))}</td><td class=small>{e(s.get('closes'))[:60]}</td></tr>")
h.append("</table><h2>5. Риски</h2><ul>" + "".join(f"<li>{e(r if isinstance(r, str) else json.dumps(r, ensure_ascii=False))[:240]}</li>" for r in O["risks"][:8]) + "</ul></body></html>")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/svet.html"
open(hp, "w").write("".join(h)); html_to_pdf(hp, "../CHELYAB_SVET_32616505090.pdf")

wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Позиции (35)"
ws.append(["№", "Наименование", "Номенкл.", "Ед.", "Кол-во", "НМЦ ед. без НДС", "НМЦ ед. с НДС", "НМЦ сумма", "Нацрежим", "Товар", "Цена ед. (с НДС)", "Цена — ориентир?", "Сумма", "НДС", "Реестр ПП 719", "Соответствие", "Источники", "Примечание"])
for r in rows:
    src = "\n".join(f"{s.get('seller','')}: {s.get('price')} — {s.get('url')}" for s in r["src"][:3] if isinstance(s, dict))
    ws.append([r["n"], r["name"], r["nom"], r["unit"], r["qty"], r["nmc_nv"], r["nmc_unit"], r["nmc_sum"], r["regime"], r["prod"], r["u"], "" if isinstance(r["up"], (int, float)) else "нет цены", r["cost"] or None, r["vat"], r["reg"], r["match"], src, r["note"]])
    for j in (6, 7, 8, 11, 13): ws.cell(ws.max_row, j).number_format = "#,##0.00"
for c in ws[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
for i, w in enumerate([4, 34, 50, 6, 7, 11, 11, 12, 14, 36, 11, 9, 12, 12, 18, 30, 50, 50], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "C2"; ws.auto_filter.ref = f"A1:R{ws.max_row}"
s3 = wb.create_sheet("Поставщики"); s3.append(["Компания", "ИНН", "Город", "Сайт", "Телефон", "E-mail", "Позиции", "Подтверждение"])
for s in O["suppliers"]: s3.append([s.get("name"), s.get("inn"), s.get("city"), s.get("site"), s.get("phone"), s.get("email"), str(s.get("closes")), str(s.get("verified"))])
s4 = wb.create_sheet("Риски"); s4.append(["Риск"]); [s4.append([(r if isinstance(r, str) else json.dumps(r, ensure_ascii=False))[:900]]) for r in O["risks"]]; s4.column_dimensions["A"].width = 140
tidy(wb, row_h=26); wb.save("../CHELYAB_SVET_32616505090_POZICII.xlsx")
print(round(C_A), round(C_B), round(C_conf), [round(net(NMC_NV * k, C)) for k in (1, .9) for C in (C_A, C_B)])
