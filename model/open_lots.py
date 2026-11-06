"""Открытые лоты ниши (ЕИС, подача с 10.11.2026) → ../OTKRYTYE_LOTY_0611.xlsx + .pdf"""
import json, html, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from pdfutil import html_to_pdf
from tidy import tidy
S = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/live/"
O = {x["purchase_number"]: x for x in json.load(open(S + "open.json"))}
DT = json.load(open(S + "det.json"))
MG = {t["number"]: t for t in json.load(open("data/companies/niche_margin.json"))["tenders"]}
reg = openpyxl.load_workbook("../NISHA_KONKURSY_DILERSTVA.xlsx").active
R0 = {str(r[0]): r for r in reg.iter_rows(min_row=2, values_only=True)}
# источник товара по ключевым словам предмета
SRC = [(r"Тулсион", "Thermax Tulsion (Индия) — бренд в ТЗ; эквивалент Токем — проверить «или эквивалент»"),
       (r"КУ ?2-8|АВ ?17-8|АН-31|катионит|анионит|ионообмен|ИОС", "«Токем» (Кемерово) / дилер «АкваХимСервис» (Москва); импорт Sunresin («Интерхим Ком»)"),
       (r"гидразин", "импорт из Китая (Yibin Tianyuan, Weifang Yaxing) — в РФ реестровых нет"),
       (r"БАУ|уголь|угля", "АО «Сорбент» (Пермь), ЭХЗ (В. Новгород)"), (r"Silcarbon", "бренд Silcarbon (Германия) — проверить эквивалент/дистрибьютора"),
       (r"силикагел", "СКАТЗ, «Уфахимкомплект»; Китай (Qingdao)"), (r"флокулянт|Гринлайф", "«Экоиндустрия» (ПАА, ПП 719), SNF China"),
       (r"полиоксихлорид|коагулянт|хлорное", "«Аурат» / «Промхимснаб», «Химпром» Кемерово"), (r"ОЭДФ|HEDP", "«Экоэнерго»; Китай (Shandong Taihe)"),
       (r"Эмикор|Эктоскейл|Оптион", "бренд в ТЗ — производитель («Экоэнерго» для Оптион/Эктоскейл); нужна дилерская цена"),
       (r"ингибитор", "«Экоэнерго», «ТехноХимРеагентБел» (ЕАЭС), Налко/CINALAB"), (r"аммиак", "«Акрон», «Новомосковский Азот», «Тольяттиазот»"),
       (r"пропиленгликол|этиленгликол", "«Химпром» Кемерово (теплоноситель), СИБУР"), (r"деэмульгатор|противотурбулент", "только реагент с ОПИ у Русвьетпетро — через производителя"),
       (r"компрессор|ЗИП", "не профиль реагентов — пропуск"), (r"водоподготовки|ВХО|водно-хим", "по спецификации (несколько реагентов)")]
import re
def src(t):
    for rx, s in SRC:
        if re.search(rx, t, re.I): return s
    return "—"
rows = []
for n in sorted(set(O) | set(DT)):
    o = O.get(n, {}); d = DT.get(n, {}); r0 = R0.get(n)
    title = o.get("object_info") or d.get("title") or (r0[3] if r0 else "")
    cust = d.get("customer") or (r0[2] if r0 else o.get("customer"))
    nmc = o.get("max_price") or d.get("max_price") or (r0[4] if r0 else None)
    close = str(o.get("submission_close_at") or d.get("close") or (r0[5] if r0 else ""))[:16]
    msp = "только МСП" if "SMBO" in str(o.get("purchase_type") or "") else ""
    m = MG.get(n, {})
    if re.search(r"КРС", title): continue          # корма для скота — не наш профиль
    skip = re.search(r"компрессор|ЗИП|деэмульгатор|противотурбулент", title, re.I)
    days = (int(close[8:10]) - 6) if close[:7] == "2026-11" else 99
    prio = "пропуск" if skip else ("A" if days >= 7 and m.get("margin_m10") else "B" if days >= 7 else "C (срок ≤6 дн.)")
    rows.append(dict(n=n, cust=cust, title=title, nmc=nmc, close=close, etp=d.get("etp") or "", msp=msp, src=src(title),
                     marg=m.get("margin_m10"), conf=m.get("confidence"), prio=prio, url=f"https://zakupki.gov.ru/epz/order/notice/notice223/common-info.html?regNumber={n}"))
rows.sort(key=lambda r: ({"A": 0, "B": 1}.get(r["prio"][0], 2) if r["prio"] != "пропуск" else 3, r["close"]))
wb = openpyxl.Workbook(); ws = wb.active; ws.title = f"Открытые лоты ({len(rows)})"
H = ["Приоритет", "Подача до", "№ ЕИС", "Заказчик", "Предмет", "НМЦ, ₽", "Площадка", "Только МСП", "Где брать товар", "Маржа при −10 % (ОСНО, если считали)", "Уверенность", "Ссылка"]
ws.append(H)
for r in rows:
    ws.append([r["prio"], r["close"], r["n"], r["cust"], r["title"], r["nmc"], r["etp"], r["msp"], r["src"], r["marg"], r["conf"], r["url"]])
    X = ws.max_row; ws.cell(X, 6).number_format = "#,##0"; ws.cell(X, 10).number_format = "0%"; ws.cell(X, 12).hyperlink = r["url"]
    col = {"A": "c6efce", "B": "fff2cc", "C": "f2f2f2"}.get(r["prio"][0], "f8cbad")
    for c in ws[X]: c.fill = PatternFill("solid", fgColor=col)
for c in ws[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
for i, w in enumerate([12, 16, 13, 26, 55, 13, 28, 11, 45, 12, 11, 40], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "C2"; ws.auto_filter.ref = f"A1:L{ws.max_row}"
plat = {}
for r in rows:
    if r["prio"] != "пропуск": plat.setdefault(r["etp"] or "уточнить", []).append(r["n"])
s2 = wb.create_sheet("Площадки"); s2.append(["Площадка", "Лотов", "Номера"])
for k, v in sorted(plat.items(), key=lambda kv: -len(kv[1])): s2.append([k, len(v), ", ".join(v)])
s2.column_dimensions["A"].width = 60; s2.column_dimensions["C"].width = 120
tidy(wb, row_h=30); wb.save("../OTKRYTYE_LOTY_0611.xlsx")

e = html.escape; fm = lambda v: f"{v:,.0f}".replace(",", " ") if isinstance(v, (int, float)) else "—"
css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:7.9px;color:#111;margin:0}h1{font-size:13px;margin:0 0 3px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #bbb;padding:1px 3px;vertical-align:top}
th{background:#305496;color:#fff;text-align:left}td.n{text-align:right;white-space:nowrap}.A{background:#c6efce}.B{background:#fff2cc}.C{background:#f2f2f2}.p{background:#f8cbad}.box{border:1px solid #305496;padding:4px 6px;background:#f3f6fb;margin:4px 0}"""
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body><h1>Открытые закупки по тематике (ЕИС, подача с 10.11.2026) — {len(rows)} лотов на 06.11.2026</h1>",
     "<div class='box'><b>A</b> — срок ≥7 дней и маржа уже посчитана; <b>B</b> — срок ≥7 дней, нужен расчёт/КП; <b>C</b> — срок ≤6 дней (успеть, только если уже есть аккредитация и товар); <b>пропуск</b> — не профиль или нужен реагент с ОПИ. "
     "Маржа — для компании на ОСНО (заказчики энергетики сравнивают цены без НДС; на УСН минус ≈18 п.п.). Главная площадка — <b>B2B-Center</b>: там почти все лоты Т Плюс. Лоты закрытых площадок нефтяников сюда не входят.</div>",
     "<table><tr><th>Пр.</th><th>Подача до</th><th>№</th><th>Заказчик</th><th>Предмет</th><th>НМЦ</th><th>Площадка</th><th>МСП</th><th>Где брать товар</th><th>Маржа −10 %</th></tr>"]
for r in rows:
    k = {"A": "A", "B": "B", "C": "C"}.get(r["prio"][0], "p")
    h.append(f"<tr class='{k}'><td>{e(r['prio'])}</td><td>{e(r['close'])}</td><td>{r['n']}</td><td>{e(str(r['cust']))[:30]}</td><td>{e(r['title'])[:95]}</td><td class=n>{fm(r['nmc'])}</td><td>{e(r['etp'])[:28]}</td><td>{e(r['msp'])}</td><td>{e(r['src'])[:70]}</td><td class=n>{'' if r['marg'] is None else f'{r[chr(109)+chr(97)+chr(114)+chr(103)]*100:.0f} %'}</td></tr>")
h.append("</table></body></html>")
hp = S + "open.html"; open(hp, "w").write("".join(h)); html_to_pdf(hp, "../OTKRYTYE_LOTY_0611.pdf")
import collections; print(len(rows), collections.Counter(r["prio"][0] for r in rows), {k: len(v) for k, v in plat.items()})
