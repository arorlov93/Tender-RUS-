"""ПАО «Россети Северный Кавказ» 32616499234 — рамочный договор линейной арматуры: расчёт по 927 позициям → ../ROSSETI_SK_32616499234.pdf
Данные: data/zakupki_0211/tz/deep/rsk_{A,B,C,D}_out.json (цены рынка, агенты 03–04.11.2026), rsk_items.json (прайс заказчика)."""
import json, html, os
from pdfutil import html_to_pdf

D = "data/zakupki_0211/tz/deep"
G = {"A": "Гасители вибрации", "B": "Зажимы (вкл. монтажный инструмент МК/МКЗ)", "C": "Сцепная и защитная арматура (звенья, распорки, рога, узлы…)", "D": "Балласты, шары-маркеры, грузы"}
ALL = {x["n"]: x for x in json.load(open(f"{D}/rsk_items.json"))}
OUT = {g: json.load(open(f"{D}/rsk_{g}_out.json")) for g in G if os.path.exists(f"{D}/rsk_{g}_out.json")}
TERMS = OUT.get("D", {}).get("terms", {})
LBL = {"criteria": "Критерии оценки", "price_formula": "Как считается цена заявки", "unit_price_rule": "Единичные расценки на переторжке", "anomaly_low": "Аномально низкая цена", "retender": "Переторжка", "participant_requirements": "Требования к участнику", "participant": "Требования к участнику", "security": "Обеспечение исполнения", "contract_security": "Обеспечение исполнения", "monitoring": "Мониторинг цен по договору", "penalties": "Неустойки", "termination": "Расторжение", "warranty": "Гарантия", "term": "Срок договора", "delivery": "Поставка", "volume": "Объём", "payment": "Оплата", "docs": "Документы в заявку"}
NMC_NET, NMC = 40983606.56, 50000000.0
EXCL = {349}  # ошибка НМЦ ×10 (МК-2 кл.2 — 203 894 при кл.1 20 390)
f = lambda v: f"{v:,.0f}".replace(",", " ") if isinstance(v, (int, float)) else (v if v not in (None, "") else "—")
e = lambda s: html.escape(str(s if s is not None else "—"))
pct = lambda v: f"{v*100:.1f} %" if v is not None else "—"
GRP = {}
for g, d in OUT.items():
    for x in d["items"]: GRP[x["n"]] = g
rows_g = []; tot_n = tot_m = tot_all = 0; cnt_p = 0
for g, name in G.items():
    d = OUT.get(g)
    if not d: rows_g.append((name, 0, 0, None, None, None, None, "не посчитано")); continue
    it = d["items"]; pr = [x for x in it if x.get("price") and x["n"] not in EXCL]
    sn = sum(x["nmc_unit"] for x in pr); sm = sum(x["price"] for x in pr); sa = sum(x["nmc_unit"] for x in it)
    tot_n += sn; tot_m += sm; tot_all += sa; cnt_p += len(pr)
    rows_g.append((name, len(it), len(pr), sa, sn, sm, (sn - sm) / sn if sn else None, ""))
mg = (tot_n - tot_m) / tot_n if tot_n else 0
css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:8.4px;color:#111;margin:18px}h1{font-size:16px;margin:0 0 4px}
h2{font-size:12.5px;margin:7px 0 3px;border-bottom:2px solid #305496;padding-bottom:2px}table{border-collapse:collapse;width:100%;margin:3px 0 8px}
th,td{border:1px solid #bbb;padding:1px 3px;vertical-align:top}th{background:#305496;color:#fff;text-align:left}td.n{text-align:right;white-space:nowrap}
.ok{background:#c6efce}.bad{background:#f8cbad}.warn{background:#fff2cc}.small{font-size:8.5px;color:#444}.box{border:1px solid #305496;padding:6px 8px;background:#f3f6fb;margin:6px 0}"""
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body>",
     "<h1>ПАО «Россети Северный Кавказ» — рамочный договор поставки линейной арматуры, № 32616499234</h1>",
     "<div class='small'>ООО «ГИДРОПАС» (УСН). Расчёт 03–04.11.2026 по прайсу заказчика (Приложение 1 к ТЗ, 927 позиций, предельные цены за ед. с НДС) и открытым ценам дилеров/заводов. Все цены рынка — с НДС, если не помечено «?». Подтверждения КП нет.</div>",
     "<h2>1. Условия</h2><table>"]
for k, v in [("Способ", "Конкурс в электронной форме (223-ФЗ), rosseti.roseltorg.ru; изменения от 03.11.2026"), ("Подача / рассмотрение / переторжка / итоги", "12.11.2026 12:00 МСК / 26.11 / 30.11 12:00–01.12 12:00 (обязательная, очная, шаг 0,5 %) / 02.12"),
             ("НМЦ (предельная сумма договора на год)", f"{f(NMC_NET)} ₽ без НДС ({f(NMC)} ₽ с НДС); договор заключается на полную НМЦ, объём НЕ гарантирован"),
             ("Сумма единичных расценок заказчика (по 1 шт. каждой поз.)", "7 789 683,53 ₽ с НДС / 6 384 986,50 ₽ без НДС — по ней сравнивают цены участников"),
             ("Поставка", "по заявкам (уведомление за 30 дней), мин. партия 50 тыс. ₽, 6 филиалов: Махачкала, Назрань, Нальчик, Черкесск, Владикавказ, Пятигорск — доставка поставщиком"),
             ("Оплата", "7 рабочих дней (МСП), аванса нет"), ("Требования к товару", "ГОСТ Р 51177-2017, УХЛ1, гарантия 5 лет, срок службы 40 лет, эквиваленты допускаются"),
             ("ПП 1875", "«преимущество» по всем 927 позициям (заявка только с российским товаром −15 % при оценке)")] + [("Оценка", "цена 70 % (сумма ед. расценок без НДС, Smin/Si); аттестация Россети на пневм. гасители 15 %; на натяжную/поддерживающую/сцепную арматуру 15 % (100/0)"), ("Переторжка", "обязательная, многократная, шаг 0,5 %; ед. расценки выше поданных — нельзя; снижение ≥25 % — доп. обеспечение")]:
    h.append(f"<tr><th style='width:20%'>{e(k)}</th><td>{e(v)}</td></tr>")
h.append("</table><h2>2. Итог по группам (по позициям, где найдена цена рынка; п.349 исключён как ошибка НМЦ)</h2>"
         "<table><tr><th>Группа</th><th>Позиций</th><th>С ценой</th><th>Σ НМЦ ед. по группе</th><th>Σ НМЦ ед. (найденные)</th><th>Σ рынок (найденные)</th><th>Маржа</th><th>Примечание</th></tr>")
for r in rows_g:
    cls = "ok" if (r[6] or 0) >= 0.3 else "warn" if (r[6] or 0) > 0 else "bad"
    h.append(f"<tr class='{cls}'><td>{e(r[0])}</td><td class=n>{r[1]}</td><td class=n>{r[2]}</td><td class=n>{f(r[3])}</td><td class=n>{f(r[4])}</td><td class=n>{f(r[5])}</td><td class=n>{pct(r[6])}</td><td>{e(r[7])}</td></tr>")
h.append(f"<tr><th>Итого</th><th class=n>{sum(r[1] for r in rows_g)}</th><th class=n>{cnt_p}</th><th class=n>{f(tot_all)}</th><th class=n>{f(tot_n)}</th><th class=n>{f(tot_m)}</th><th class=n>{pct(mg)}</th><th>покрытие {pct(tot_n/tot_all if tot_all else 0)} суммы прайса</th></tr></table>")
h.append("<h2>3. Сценарии (на прайс-листе; договор по заявкам до 40,98 млн без НДС)</h2><table><tr><th>Снижение суммы ед. расценок</th><th>Маржа к рынку</th><th>Выручка при выборке всей НМЦ, ₽</th><th>Валовая прибыль при выборке всей НМЦ, ₽</th><th>Чистая (−0,8 % ЭТП, доставка 3 %, 20 % годовых × 60 дн., УСН 15 %), ₽</th></tr>")
for drop in (0, 0.10, 0.15, 0.20, 0.25):
    m = 1 - (1 - mg) / (1 - drop); rev = NMC_NET  # договор на полную НМЦ — наша цена ед. ниже, объём в деньгах тот же
    cost = rev * (1 - m); gross = rev - cost; pre = gross - 0.008 * rev - 0.03 * rev - cost * 0.2 * 60 / 365
    cls = "ok" if m >= 0.25 else "warn" if m > 0.1 else "bad"
    h.append(f"<tr class='{cls}'><td>−{int(drop*100)} %{' (порог аномально низкой цены — доп. обеспечение)' if drop>=0.25 else ''}</td><td class=n>{pct(m)}</td><td class=n>{f(rev)}</td><td class=n>{f(gross)}</td><td class=n>{f(pre*0.85)}</td></tr>")
h.append("</table><div class='small'>Важно: объём не гарантирован — заказчик выбирает по заявкам; расчёт «при выборке всей НМЦ» — верхняя граница. Маржа взята средневзвешенной по найденным позициям; ассортимент реальных заявок может сместить её.</div>")
ab = sorted({n for d in OUT.values() for n in (d.get("above_nmc") or [])})
h.append("<h2>4. Позиции, где рынок дороже цены заказчика</h2><p>" + (", ".join(f"№{n} {e(ALL[n]['name'])}" for n in ab if n in ALL) or "нет") + "</p>")
h.append("<h2>5. Изготовители и дилеры для КП</h2><table><tr><th>Группа</th><th>Компания</th><th>Телефон</th><th>E-mail</th><th>Сайт</th><th>ИНН</th><th>2 источника</th></tr>")
seen = set()
for g, d in OUT.items():
    for m in d.get("makers", []):
        k = (m.get("name") or "")[:40]
        if k in seen or len(seen) >= 7: continue
        seen.add(k)
        h.append(f"<tr><td>{g}</td><td>{e(m.get('name'))}</td><td>{e(str(m.get("phone") or "").split(";")[0])}</td><td>{e(str(m.get("email") or "").split(",")[0])}</td><td>{e(m.get('site'))}</td><td>{e(m.get('inn'))}</td><td>{'да' if m.get('verified') else 'нет'}</td></tr>")
h.append("</table><h2>6. Риски</h2><ul>")
for r in ["Аттестация ПАО «Россети» на пневматические гасители и на натяжную/поддерживающую/сцепную арматуру — 30 % оценки (100/0 баллов). Нужны заключения изготовителей (МЗВА и др.) и письма о полномочиях.",
          "Обеспечение исполнения: в Информационной карте «не предусмотрено», в проекте договора п.4.4 — 5 % НМЦ (≈2,5 млн) — вопрос до 09.11 12:00.",
          "Мониторинг цен (п.6.9–6.11 договора): заказчик вправе требовать снизить цены до минимальных рыночных; при отказе — расторжение. Право одностороннего отказа без компенсации (п.15.6).",
          "Неустойки 0,2 %/день; штраф 10 % от цены договора за документы о стране происхождения.",
          "Переторжка многократная с шагом 0,5 %; снижение ≥25 % — аномально низкая цена → доп. обеспечение.",
          "Позиция №349 МК-2 кл.2: НМЦ 203 894 ₽ — вероятная ошибка ×10 (кл.1 — 20 390 ₽); МК/МКЗ — монтажный инструмент, не линейная арматура.",
          "Цены рынка — розница дилеров и прайсы без даты; часть позиций без цены (по запросу у заводов). До подачи нужны КП МЗВА/Форэнерго-Трейд, ТАИЗ, ТЭМЗ, ЮАИЗ, ЭССП."]:
    h.append(f"<li>{e(r)}</li>")
h.append("</ul>")
h.append("<h2>7. По группам: крупнейшие позиции (по НМЦ ед.)</h2>")
for g, name in G.items():
    d = OUT.get(g)
    if not d: continue
    top = sorted(d["items"], key=lambda z: -(ALL.get(z["n"], {}).get("nmc_unit") or 0))[:6]
    h.append(f"<b>{g}. {e(name)}</b><table><tr><th>№</th><th>Позиция</th><th>НМЦ ед.</th><th>Рынок</th><th>Маржа</th><th>Товар / изготовитель</th></tr>")
    for x in top:
        a = ALL[x["n"]]["nmc_unit"]; p = x.get("price"); m = (a - p) / a if p else None
        cls = "" if m is None else "ok" if m >= 0.3 else "warn" if m >= 0 else "bad"
        h.append(f"<tr class='{cls}'><td>{x['n']}</td><td>{e(ALL[x['n']]['name'])[:70]}</td><td class=n>{f(a)}</td><td class=n>{f(p)}</td><td class=n>{pct(m)}</td><td>{e((x.get('maker') or x.get('product') or x.get('match') or ''))[:90]}</td></tr>")
    h.append("</table>")
h.append("</body></html>")
open("/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/rsk.html", "w").write("".join(h))
html_to_pdf("/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/rsk.html", "../ROSSETI_SK_32616499234.pdf")
print("групп", len(OUT), "с ценой", cnt_p, "НМЦ найд", f(tot_n), "рынок", f(tot_m), "маржа", pct(mg))

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from tidy import tidy
wb = openpyxl.Workbook(); s0 = wb.active; s0.title = "Итог"
s0.append(["Россети СК 32616499234 — линейная арматура: итог по группам (рынок с НДС против предельной цены ед. заказчика с НДС)"]); s0["A1"].font = Font(bold=True, size=12)
s0.append(["Группа", "Позиций", "С ценой рынка", "Σ НМЦ ед. по группе", "Σ НМЦ ед. (найденные)", "Σ рынок (найденные)", "Маржа", "Примечание"])
for c in s0[2]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")
for r in rows_g: s0.append(list(r))
s0.append(["ИТОГО", sum(r[1] for r in rows_g), cnt_p, tot_all, tot_n, tot_m, mg, "п.349 исключён (ошибка НМЦ ×10)"])
for row in s0.iter_rows(min_row=3):
    for j in (3, 4, 5): row[j].number_format = "#,##0"
    row[6].number_format = "0.0%"
for i, w in enumerate([44, 9, 11, 16, 16, 16, 9, 40], 1): s0.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
for g, name in G.items():
    if g not in OUT: continue
    ws = wb.create_sheet(f"{g} {name}"[:31])
    ws.append(["№", "Позиция (прайс заказчика)", "НМЦ ед. с НДС", "Рынок ед.", "НДС в цене рынка", "Маржа ед.", "Товар (продавец)", "Изготовитель", "Совпадение / примечание", "Ссылка"])
    for c in ws[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
    for x in sorted(OUT[g]["items"], key=lambda z: z["n"]):
        p = x.get("price"); a = ALL.get(x["n"], {}).get("nmc_unit", x.get("nmc_unit"))
        m = (a - p) / a if p and a else None
        ws.append([x["n"], ALL.get(x["n"], {}).get("name", x.get("name")), a, p, str(x.get("vat")), m, x.get("product"), x.get("maker"), x.get("match"), x.get("url")])
        X = ws.max_row; ws.cell(X, 3).number_format = "#,##0.00"; ws.cell(X, 4).number_format = "#,##0.00"; ws.cell(X, 6).number_format = "0%"
        if m is not None: ws.cell(X, 6).fill = PatternFill("solid", fgColor="C6EFCE" if m >= 0.3 else "FFF2CC" if m >= 0 else "F8CBAD")
        if x.get("url"): ws.cell(X, 10).hyperlink = x["url"]
    for i, w in enumerate([6, 46, 13, 12, 9, 8, 40, 26, 40, 36], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = "C2"; ws.auto_filter.ref = f"A1:J{ws.max_row}"
mk = wb.create_sheet("Изготовители")
mk.append(["Группа", "Компания", "Телефон", "E-mail", "Сайт", "ИНН", "2 источника"])
for g, d in OUT.items():
    for m in d.get("makers", []): mk.append([g, m.get("name"), m.get("phone"), m.get("email"), m.get("site"), m.get("inn"), "да" if m.get("verified") else "нет"])
tidy(wb, row_h=24); wb.save("../ROSSETI_SK_32616499234_POZICII.xlsx")
