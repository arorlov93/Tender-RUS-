"""Итог по ТЗ: маржа 6 лотов после чтения ТЗ + перепроверка. data/companies/prices_tz.json → ../ITOG_MARZHA_PO_TZ_0611.xlsx + .pdf"""
import json, ast, html, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from pdfutil import html_to_pdf
from tidy import tidy
D = json.load(open("data/companies/prices_tz.json"))
P = lambda v: ast.literal_eval(v) if isinstance(v, str) and v[:1] in "{[" else v
def T(v):
    v = P(v)
    if isinstance(v, dict): return "; ".join(f"{k}: {T(x)}" for k, x in v.items())
    if isinstance(v, list): return "; ".join(T(x) for x in v)
    return "" if v in (None, "None") else str(v)
num = lambda v: float(v) if v not in (None, "", "None") else None
# сверка с ТЗ (проверено вручную по файлам data/tz_lots, 06.11.2026)
CHECK = {
 "32616534213": ("подаём при КП завода", "Сверено: эквивалент допускается, но по ТЗ — только натриевые соли НТФ ≥25 % и pH 5,5–7,5 (ОЭДФ не проходит); количество в пересчёте на 100 % (×4 товарного); НМЦ без НДС, оценка без НДС; ограничение (ПП 719 — Нитон, Эколарис)."),
 "32616534438": ("подаём при КП ≤1 000 ₽/кг", "Сверено: эквивалент допускается (эмульсия ОДА, ОДА ≥94 % на сухое); 3 адреса, 3 партии (30 дней / 09.04.2027 / 16.07.2027); НМЦ без НДС; МСП."),
 "32616531631": ("подаём с эквивалентом", "Сверено: K814, эквивалент допускается (кокосовый, 1,2–2,5 мм ≥90 %, йод ≥1050, BET ≥1100); преимущество (импорт допустим); до 11.05.2027; НМЦ без НДС."),
 "32616548235": ("подаём при КП ≤750 тыс./т", "Сверено: 16,3 т (5,7 Владимир + 10,6 Иваново), апрель 2027, эквивалент НЕ допускается, обеспечений нет. ВНИМАНИЕ: формула оценки в «Критериях» — стоимость без НДС (на B2B-Center указано «с НДС») — для УСН считать худший случай (без НДС), задать вопрос заказчику."),
 "32616522434": ("не идём (только при КП ≤17,5 тыс./т)", "Сверено: марка А ГОСТ 9-92, 700 т, партии 22±3 т за 4 дня, 2027 г.; НМЦ 21 122 500 «включая все налоги» (с НДС) — маржа ниже прежней оценки; аукцион; ограничение (марка А в ПП 719 — Уралхим, АМ Сервис)."),
 "32616539091": ("не идём без кредитной линии", "Сверено: ПОХА 17,3 % Al2O3, 2 100 т, период ДО 31.03.2027 (≈525 т/мес), партия — 22 раб. дня; оплата ЧЕРЕЗ 90 КАЛЕНДАРНЫХ ДНЕЙ после поставки → оборотка 75–100 млн; питьевая вода (СанПиН — экспертное заключение); обеспечения 1,28 + 6,4 млн."),
}
rows = []
for l in D["lots"]:
    m = P(l.get("margin")) or {}; pr = P(l.get("profit_rub")) or {}; fl = P(l.get("floor_price")) or {}
    n = str(l["number"]); dec, chk = CHECK.get(n, ("", ""))
    if n == "32616548235": m = dict(m, usn_nmc=0.165, usn_m10=0.073, usn_m20=-0.043)   # худший случай: оценка без НДС
    rows.append(dict(n=n, cust=T(l.get("customer")), product=T(l.get("product"))[:120], qty=T(l.get("qty")), unit=T(l.get("unit")), nmc=num(l.get("nmc")), buy=T(l.get("best_buy_unit")),
                     buy_note=T(l.get("best_buy_note"))[:300], cost=num(l.get("cost_vat")), m=m, pr=pr, fl=fl, wc=T(l.get("working_capital"))[:300], sec=T(l.get("security"))[:300],
                     vat=T(l.get("vat_basis"))[:200], risks=T(l.get("risks"))[:600], verdict=T(l.get("verdict"))[:400], conf=T(l.get("confidence")), dec=dec, chk=chk, sched=T(l.get("schedule"))[:400], offers=P(l.get("offers")) or []))
order = {"32616531631": 1, "32616548235": 2, "32616534438": 3, "32616534213": 4, "32616539091": 5, "32616522434": 6}
rows.sort(key=lambda r: order.get(r["n"], 9))
cls = lambda v: "c6efce" if isinstance(v, (int, float)) and v >= .15 else "fff2cc" if isinstance(v, (int, float)) and v > 0 else "f8cbad"
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Маржа по ТЗ (6 лотов)"
H = ["Решение", "№ ЕИС", "Заказчик", "Товар", "Кол-во", "НМЦ, ₽", "Закупка за ед. (с НДС)", "Себестоимость с НДС", "УСН НМЦ", "УСН −10 %", "УСН −20 %", "ОСНО НМЦ", "ОСНО −10 %", "ОСНО −20 %",
     "Прибыль при −10 %, УСН", "Прибыль при −10 %, ОСНО", "Минимальная цена заявки (10 % маржи)", "Сверка с ТЗ", "График поставки", "Сравнение НДС", "Оборотка", "Обеспечение", "Риски", "Комментарий к закупочной цене", "Уверенность"]
ws.append(H)
for r in rows:
    m = r["m"]
    ws.append([r["dec"], r["n"], r["cust"], r["product"], f"{r['qty']} {r['unit']}", r["nmc"], r["buy"], r["cost"], m.get("usn_nmc"), m.get("usn_m10"), m.get("usn_m20"), m.get("osno_nmc"), m.get("osno_m10"), m.get("osno_m20"),
               num(r["pr"].get("usn_m10")), num(r["pr"].get("osno_m10")), T(r["fl"]), r["chk"], r["sched"], r["vat"], r["wc"], r["sec"], r["risks"], r["buy_note"], r["conf"]])
    X = ws.max_row
    for j in (6, 8, 15, 16): ws.cell(X, j).number_format = "#,##0"
    for j in range(9, 15):
        ws.cell(X, j).number_format = "0.0%"; ws.cell(X, j).fill = PatternFill("solid", fgColor=cls(ws.cell(X, j).value))
for c in ws[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
for i, w in enumerate([22, 13, 22, 34, 14, 13, 14, 14, 9, 9, 9, 9, 9, 9, 13, 13, 34, 60, 45, 30, 40, 40, 60, 50, 10], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "E2"
s2 = wb.create_sheet("Предложения продавцов"); s2.append(["№ ЕИС", "Продавец", "Товар", "Цена за ед. с НДС", "База", "Дата", "Ссылка"])
for r in rows:
    for o in r["offers"]:
        if isinstance(o, dict): s2.append([r["n"], o.get("seller"), T(o.get("product"))[:100], num(o.get("price_unit_vat")) if str(o.get("price_unit_vat")).replace(".", "").isdigit() else o.get("price_unit_vat"), o.get("basis"), o.get("date"), o.get("url")])
for c in s2[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")
for i in range(1, 8): s2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 30
tidy(wb, row_h=60); wb.save("../ITOG_MARZHA_PO_TZ_0611.xlsx")
e = html.escape; fm = lambda v: f"{v:,.0f}".replace(",", " ") if isinstance(v, (int, float)) else "—"; pc = lambda v: f"{v*100:.0f} %" if isinstance(v, (int, float)) else "—"
css = "body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:7.8px;margin:0}h1{font-size:13px;margin:0 0 3px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #bbb;padding:2px 3px;vertical-align:top}th{background:#305496;color:#fff}td.n{text-align:right;white-space:nowrap}.box{border:1px solid #305496;padding:4px 6px;background:#f3f6fb;margin:4px 0}"
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body><h1>Маржа по ТЗ — 6 лотов (сверено с документацией, 06.11.2026)</h1>",
     "<div class='box'>Подтверждённых КП заводов пока нет — закупочные цены: цены перепродавцов или допущения (см. Excel). Маржа валовая (без налога, процентов, стоимости гарантии). УСН — ГИДРОПАС сейчас; ОСНО — новая компания. "
     "<b>Налоговый риск УСН:</b> при доходе за 2027 г. свыше 15 млн ГИДРОПАС становится плательщиком НДС 5 % — маржа УСН падает ещё ≈4,8 п.п.; при двух и более победах выгоднее ОСНО.</div>",
     "<table><tr><th>Решение</th><th>Лот</th><th>НМЦ</th><th>Себест.</th><th>УСН НМЦ</th><th>УСН −10%</th><th>УСН −20%</th><th>ОСНО −10%</th><th>Прибыль −10% УСН</th><th>Сверка с ТЗ</th></tr>"]
for r in rows:
    m = r["m"]
    h.append(f"<tr><td><b>{e(r['dec'])}</b></td><td>{r['n']}<br>{e(r['product'][:60])}</td><td class=n>{fm(r['nmc'])}</td><td class=n>{fm(r['cost'])}</td>" +
             "".join(f"<td class=n style='background:#{cls(m.get(k))}'>{pc(m.get(k))}</td>" for k in ("usn_nmc", "usn_m10", "usn_m20", "osno_m10")) +
             f"<td class=n>{fm(num(r['pr'].get('usn_m10')))}</td><td>{e(r['chk'])}</td></tr>")
h.append("</table></body></html>")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/itz.html"; open(hp, "w").write("".join(h)); html_to_pdf(hp, "../ITOG_MARZHA_PO_TZ_0611.pdf"); print("ok")
