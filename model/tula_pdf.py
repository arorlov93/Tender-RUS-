"""ПАО «Россети Центр и Приволжье» — Тулэнерго, конкурс МСП 32616505169 «Электроизоляционные материалы» (лот 402A), единичные расценки.
Данные: deep/tula_items.json (НМЦ ед. и справочник — со скана расчёта НМЦ), deep/tula_out.json (рынок, агент 04.11.2026). → ../TULA_32616505169.pdf + _POZICII.xlsx"""
import json, html, statistics, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from pdfutil import html_to_pdf
from tidy import tidy

D = "data/zakupki_0211/tz/deep/"
IT = {x["n"]: x for x in json.load(open(D + "tula_items.json"))}
O = json.load(open(D + "tula_out.json")); R = {int(x["n"]): x for x in O["items"]}
LIMIT = 4058188.48
num = lambda v: float(str(v).replace(",", ".").replace(" ", "")) if v not in (None, "", "None") else None
f = lambda v: f"{v:,.0f}".replace(",", " ") if isinstance(v, (int, float)) else "—"
f2 = lambda v: f"{v:,.2f}".replace(",", " ") if isinstance(v, (int, float)) else "—"
e = lambda s: html.escape(str(s if s not in (None, "None") else "—"))
MARK = 1.25   # наша цена = рынок × 1,25 (маржа 20 %), но не выше НМЦ ед.
rows = []
for n in sorted(IT):
    x = IT[n]; r = R.get(n, {}); m = num(r.get("unit_price"))
    bid = min(x["nmc_unit"], m * MARK) if m else x["nmc_unit"]
    rows.append(dict(x, m=m, bid=bid, prod=r.get("product"), reg=r.get("registry"), match=r.get("match"), note=r.get("note"), src=r.get("sources") or [], vat=r.get("vat"), origin=r.get("origin"), unit_m=r.get("unit")))
S_nmc = sum(r["nmc_unit"] for r in rows); S_ref = sum(r["ref_unit"] for r in rows); S_m = sum(r["m"] for r in rows if r["m"]); S_bid = sum(r["bid"] for r in rows)
marg = [1 - r["m"] / r["nmc_unit"] for r in rows if r["m"]]
over = [r for r in rows if r["m"] and r["m"] > r["nmc_unit"]]
under_ref = sum(1 for r in rows if r["m"] and r["m"] > r["ref_unit"])
# маржа по объёму договора при выборке, пропорциональной НМЦ ед. (один «комплект» всех позиций)
m_nmc = 1 - S_m / S_nmc; m_bid = 1 - S_m / S_bid
EX = [r for r in rows if r["n"] != 42]   # поз. 42 — вероятная ошибка ед. изм. в НМЦ (420 ₽/кг при рынке ≈3 600)
X_m = sum(r["m"] for r in EX); X_nmc = sum(r["nmc_unit"] for r in EX); X_bid = sum(r["bid"] for r in EX); X_ref = sum(r["ref_unit"] for r in EX)
xm_nmc = 1 - X_m / X_nmc; xm_bid = 1 - X_m / X_bid
def net(P, C, days=40):
    pre = P - C - 0.008 * P - C * 0.2 * days / 365 - 0.03 * C   # доставка партиями по заявкам ≈3 % закупки
    return pre - 0.15 * max(pre, 0)
css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:8.2px;color:#111;margin:0}h1{font-size:14px;margin:0 0 3px}
h2{font-size:11px;margin:6px 0 3px;border-bottom:2px solid #305496;padding-bottom:1px}table{border-collapse:collapse;width:100%;margin:2px 0 4px}
th,td{border:1px solid #bbb;padding:1px 3px;vertical-align:top}th{background:#305496;color:#fff;text-align:left}td.n{text-align:right;white-space:nowrap}
.ok{background:#c6efce}.bad{background:#f8cbad}.warn{background:#fff2cc}.small{font-size:7.1px;color:#444}.box{border:1px solid #305496;padding:4px 6px;background:#f3f6fb;margin:4px 0}"""
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body>",
     "<h1>ПАО «Россети Центр и Приволжье» — «Тулэнерго»: электроизоляционные материалы (лот 402A), ЕИС 32616505169</h1>",
     "<div class='small'>ООО «ГИДРОПАС» (УСН). Извещение, ТЗ + приложение (70 поз.), критерии, проект договора, расчёт НМЦ (скан). Рынок — открытые источники на 04.11.2026, с НДС, без доставки; КП не запрашивались. Все ссылки — в Excel.</div><h2>1. Условия</h2><table>"]
for a, b in [("Цена", f"договор по ЕДИНИЧНЫМ РАСЦЕНКАМ: лимит {f2(LIMIT)} ₽ с НДС, сумма цен ед. 70 поз. = {f2(S_nmc)} ₽; объёмы не заданы — по заявкам до 31.12.2027"),
             ("Подача", "конкурс МСП, РАД (lot-online); до 09.11.2026 10:00 МСК; 1-е части до 17.11, 2-е до 23.11, доп. ценовые предложения с 24.11"),
             ("Оценка", "цена 70 баллов — ПО КАЖДОЙ позиции (мин. цена / наша цена, все 70 поз. равновесны); опыт 15; гарантия 10; срок поставки 5. УСН — цена = без НДС, сравнивается с ценами с НДС"),
             ("Поставка", "10 календарных дней с заявки, авто с верхней разгрузкой, 5 складов: Тула (2), Новомосковск, Суворов, Ефремов"),
             ("Оплата / пени", "7 раб. дней после ТОРГ-12/УПД; пени 0,2 %/день; обеспечения нет (кроме демпинга −25 %)"),
             ("Нацрежим", "изоленты и ленты (поз. 1–23, 26–31) — ограничение; остальные — преимущество 15 %")]:
    h.append(f"<tr><th style='width:10%'>{a}</th><td>{e(b)}</td></tr>")
h.append(f"</table><h2>2. Сводка: сумма цен единиц</h2><table><tr><th>НМЦ</th><th>Справочник заказчика</th><th>Рынок (лучшие найденные)</th><th>Рынок / НМЦ</th><th>Рынок / справочник</th><th>Поз. рынок > НМЦ</th><th>Поз. рынок > справочника</th><th>Медиана маржи поз. по НМЦ</th></tr>"
         f"<tr><td class=n>{f2(S_nmc)}</td><td class=n>{f2(S_ref)}</td><td class=n>{f2(S_m)}</td><td class=n>{S_m/S_nmc*100:.0f} %</td><td class=n>{S_m/S_ref*100:.0f} %</td><td class=n>{len(over)}</td><td class=n>{under_ref}</td><td class=n>{statistics.median(marg)*100:.0f} %</td></tr></table>")
h.append("<h2>3. Маржа на 4,06 млн (выборка пропорционально ценам ед.; без поз. 42 — см. риски)</h2><table><tr><th>Сценарий цены</th><th>Сумма цен ед.</th><th>Маржа</th><th>Чистая на весь лимит</th></tr>")
for nm, P in [("По НМЦ", X_nmc), ("Наша цена = рынок × 1,25, не выше НМЦ ед.", X_bid), ("На уровне справочника заказчика", X_ref), ("НМЦ −20 %", X_nmc * .8)]:
    k = LIMIT / P; mm = 1 - X_m / P
    h.append(f"<tr class='{'ok' if mm >= .2 else 'warn' if mm > .05 else 'bad'}'><td>{e(nm)}</td><td class=n>{f2(P)}</td><td class=n>{mm*100:.0f} %</td><td class=n>{f(net(LIMIT, X_m * k))}</td></tr>")
h.append("</table><div class='small'>Объёмы не известны: если закажут в основном позиции, где рынок дороже НМЦ (Силотерм ЭП-6К, МРПИ d50, ПВХ d32), маржа меньше. Чистая: −0,8 % ЭТП, −3 % доставка партиями, деньги 20 % год. × 40 дн., УСН 15 %; лимит выбирают редко полностью.</div>")
h.append(f"<div class='box'><b>Вывод: РИСКОВАННО — маржа есть, только если конкуренты не опустятся к «справочнику» заказчика (на его уровне маржа ≈0). По рынку запас {xm_nmc*100:.0f} % от НМЦ (без поз. 42), но победит тот, кто ближе к «справочнику» заказчика.</b> Рынок ({f(S_m)}) почти совпадает со справочником ({f(S_ref)}): конкуренты-дилеры, скорее всего, предложат около него, а оценка идёт по каждой позиции отдельно. "
         f"Рекомендация: цена = рынок × 1,2–1,3, но не выше НМЦ ед. (сумма ≈{f(X_bid)} без поз. 42, маржа ≈{xm_bid*100:.0f} %; чистая на весь лимит ≈{f(net(LIMIT, X_m * LIMIT / X_bid))} ₽), по позициям с ошибками НМЦ — НМЦ. "
         f"Риски: поз. 42 Силотерм ЭП-6К — НМЦ 420 ₽/кг при рынке ≈3 600 (вероятно, ошибка в единице) → при заказе убыток; поз. 44 текстолит — единица м² или кг (справочник 663 похож на кг); «ТЕХ 72», изолента 0,2 мм — марки не найдены; по изолентам записей ПП 719 нет ни у кого — ограничение не блокирует. "
         "Нужен запрос разъяснений до 06.11: ед. изм. поз. 42, 44, 27; размеры поз. 1; 30 или 15 шт в поз. 38. Пени 0,2 %/день при поставке за 10 дней — держать склад ходовых позиций.</div>")
h.append("<h2>4. Позиции с наибольшим отклонением</h2><table><tr><th>№</th><th>Позиция</th><th>Ед.</th><th>НМЦ</th><th>Справ.</th><th>Рынок</th><th>Маржа по НМЦ</th><th>Товар</th></tr>")
for r in sorted([r for r in rows if r["m"]], key=lambda r: r["m"] / r["nmc_unit"])[-8:][::-1] + sorted([r for r in rows if r["m"]], key=lambda r: r["m"] / r["nmc_unit"])[:6]:
    mm = 1 - r["m"] / r["nmc_unit"]
    h.append(f"<tr class='{'bad' if mm < 0 else 'ok' if mm > .3 else 'warn'}'><td>{r['n']}</td><td>{e(r['name'])}</td><td>{e(r['unit'])}</td><td class=n>{f2(r['nmc_unit'])}</td><td class=n>{f2(r['ref_unit'])}</td><td class=n>{f2(r['m'])}</td><td class=n>{mm*100:.0f} %</td><td class=small>{e(str(r['prod'])[:70])}</td></tr>")
h.append("</table><h2>5. Поставщики для КП</h2><table><tr><th>Компания</th><th>Телефон</th><th>Что закрывает</th></tr>")
for s in O["suppliers"]: h.append(f"<tr><td>{e(s.get('name'))[:70]}</td><td>{e(s.get('phone'))[:35]}</td><td class=small>{e(s.get('closes'))[:110]}</td></tr>")
h.append("</table></body></html>")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/tula.html"
open(hp, "w").write("".join(h)); html_to_pdf(hp, "../TULA_32616505169.pdf")

wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Позиции (70)"
ws.append(["№", "SAP", "Наименование", "Требования ТЗ", "Ед. ТЗ", "НМЦ ед. с НДС", "Справочник ед.", "Рынок ед.", "Ед. рынка", "Маржа по НМЦ", "Предлагаемая цена (рынок×1,25 ≤ НМЦ)", "Нацрежим", "Товар", "Происхождение", "Реестр ПП 719", "Соответствие", "НДС", "Источники", "Примечание"])
for r in rows:
    src = "\n".join(f"{s.get('seller','')}: {s.get('price')} — {s.get('url')}" for s in r["src"][:3] if isinstance(s, dict))
    ws.append([r["n"], r["sap"], r["name"], r["spec"], r["unit"], r["nmc_unit"], r["ref_unit"], r["m"], r["unit_m"], (1 - r["m"] / r["nmc_unit"]) if r["m"] else None, round(r["bid"], 2), r["regime"], r["prod"], r["origin"], None if r["reg"] in (None, "None") else r["reg"], r["match"], r["vat"], src, r["note"]])
    for j in (6, 7, 8, 11): ws.cell(ws.max_row, j).number_format = "#,##0.00"
    ws.cell(ws.max_row, 10).number_format = "0%"
ws.append([None, None, "ИТОГО сумма цен ед.", None, None, S_nmc, S_ref, S_m, None, m_nmc, round(S_bid, 2)])
for c in ws[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
for i, w in enumerate([4, 9, 30, 40, 6, 10, 10, 10, 12, 9, 12, 12, 36, 20, 16, 30, 10, 50, 50], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "D2"; ws.auto_filter.ref = f"A1:S{ws.max_row}"
s3 = wb.create_sheet("Поставщики"); s3.append(["Компания", "ИНН", "Город", "Сайт", "Телефон", "E-mail", "Позиции", "Подтверждение"])
for s in O["suppliers"]: s3.append([s.get("name"), s.get("inn"), s.get("city"), s.get("site"), s.get("phone"), s.get("email"), str(s.get("closes")), str(s.get("verified"))])
s4 = wb.create_sheet("Риски"); s4.append(["Риск"]); [s4.append([(r if isinstance(r, str) else json.dumps(r, ensure_ascii=False))[:900]]) for r in O["risks"]]; s4.column_dimensions["A"].width = 140
tidy(wb, row_h=26); wb.save("../TULA_32616505169_POZICII.xlsx")
print(round(S_nmc), round(S_ref), round(S_m), round(S_bid), round(m_nmc, 3), round(m_bid, 3), len(over), round(net(LIMIT, S_m * LIMIT / S_bid)))
