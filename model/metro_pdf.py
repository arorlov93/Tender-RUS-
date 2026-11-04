"""ГУП «Петербургский метрополитен» 32616525105 — электроустановочные материалы 11 (ЗК МСП, 11 поз.).
Данные: data/zakupki_0211/tz/deep/metro_items.json (НМЦ ед. из ЕИС), metro_out.json (+ metro_out2.json, если есть — добор).
→ ../METRO_SPB_32616525105.pdf (2 стр.) + ../METRO_SPB_32616525105_POZICII.xlsx"""
import json, html, os, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from pdfutil import html_to_pdf
from tidy import tidy

D = "data/zakupki_0211/tz/deep/"
IT = json.load(open(D + "metro_items.json"))
O = {}; SUP = []; RISK = []
for fn in ("metro_out.json", "metro_out2.json", "metro_out3.json"):
    if not os.path.exists(D + fn): continue
    d = json.load(open(D + fn))
    for x in d["items"]:
        if fn == "metro_out.json" or x.get("unit_price") or int(x["n"]) not in O: O[int(x["n"])] = x
    SUP += d.get("suppliers", []); RISK += [r if isinstance(r, str) else json.dumps(r, ensure_ascii=False) for r in d.get("risks", [])]
NMC = sum(x["nmc_unit"] for x in IT)   # 37 398,13 — сумма цен единиц, по ней торг
MAXC = 1500000
f = lambda v: f"{v:,.0f}".replace(",", " ") if v is not None else "—"
f2 = lambda v: f"{v:,.2f}".replace(",", " ") if v is not None else "—"
e = lambda s: html.escape(str(s if s is not None else "—"))
num = lambda v: float(v) if isinstance(v, (int, float)) else None
rows = []
for x in IT:
    o = O.get(x["n"], {})
    rows.append(dict(x, up=num(o.get("unit_price")), prod=o.get("product") or "", match=o.get("match") or "", reg=o.get("registry") or "",
                     vat=o.get("vat") or "", src=o.get("sources") or [], note=o.get("note") or ""))
pri = [r for r in rows if r["up"]]; miss = [r for r in rows if not r["up"]]
# Мы на УСН: при снижении на s получаем по каждой позиции (1−s)·НМЦ ед. с НДС (коэф. пересчёта к цене без НДС). Маржа поз. = 1 − закупка/((1−s)·НМЦ).
DELIV = 0.03   # доставка по СПб/до склада и мелкие партии по заявкам — оценка 3 % к закупке
def marg(r, s): return 1 - r["up"] * (1 + DELIV) / ((1 - s) * r["nmc_unit"])
tgt = lambda r: 0.85 * 0.9 * r["nmc_unit"] / (1 + DELIV)   # предельная закупка для 15 % при −10 %

css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:8.6px;color:#111;margin:0}h1{font-size:14px;margin:0 0 3px}
h2{font-size:11px;margin:6px 0 3px;border-bottom:2px solid #305496;padding-bottom:1px}table{border-collapse:collapse;width:100%;margin:2px 0 4px}
th,td{border:1px solid #bbb;padding:1px 3px;vertical-align:top}th{background:#305496;color:#fff;text-align:left}td.n{text-align:right;white-space:nowrap}
.ok{background:#c6efce}.bad{background:#f8cbad}.warn{background:#fff2cc}.small{font-size:7.5px;color:#444}.box{border:1px solid #305496;padding:4px 6px;background:#f3f6fb;margin:4px 0}"""
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body>",
     "<h1>ГУП «Петербургский метрополитен» — электроустановочные материалы 11, ЕИС 32616525105</h1>",
     "<div class='small'>ООО «ГИДРОПАС» (УСН). Извещение, ТТ (Прил. 2), проект договора; НМЦ ед. — из ЕИС. Цены — открытые источники на 04.11.2026; реестр ПП 719 — выгрузка ГИСП 01.11.2026. Все ссылки — в Excel.</div>",
     "<h2>1. Условия</h2><table>"]
for k, v in [("Процедура", "запрос котировок в эл. форме, только МСП; Росэлторг (ЕЭТП); подача до 10.11.2026 12:00 МСК, итоги 11.11.2026; обеспечение заявки — нет"),
             ("Цена", f"договор рамочный: макс. {f(MAXC)} ₽ с НДС, объём — по заявкам «по потребности». Торг по сумме цен единиц 11 поз. = {f2(NMC)} ₽ с НДС; цены ед. снижаются пропорционально"),
             ("Для УСН", "цены ед. пересчитываются к НМЦ без НДС коэффициентом «наша сумма / НМЦ без НДС» — т.е. предложив ту же сумму, что плательщик НДС, получаем её целиком (НДС не платим)"),
             ("Поставка", "партиями по заявкам, 80 дней с заявки, транспортом поставщика на склады в СПб, пропуск за 2 дня; поддоны П4; гарантия ≥12 мес."),
             ("Оплата", "7 раб. дней после приёмки каждой партии, аванса нет; обеспечение исполнения 45 000 ₽ (при снижении ≥25 % — ×1,5 = 67 500 ₽), гарантия банка из перечня 44-ФЗ"),
             ("Нацрежим", "ОГРАНИЧЕНИЕ (ПП 1875) по поз. 4, 5, 9, 10 — нужен номер реестровой записи ПП 719 (или ЕАЭС) на каждую; без него заявка проигрывает/отклоняется при наличии российского предложения"),
             ("Штрафы", "некачественный товар — 10 %; контрафакт — повышенные штрафы; нарушение сроков вывоза брака — 50 000 ₽")]:
    h.append(f"<tr><th style='width:12%'>{k}</th><td>{e(v)}</td></tr>")
h.append("</table><h2>2. Позиции: НМЦ, рынок, маржа (мы получаем (1−снижение)·НМЦ ед.)</h2><table><tr><th>№</th><th>Позиция</th><th>НМЦ ед.</th><th>Доля</th><th>Рынок ед.</th><th>Маржа по НМЦ</th><th>Маржа −10 %</th><th>Предел закупки для 15 % при −10 %</th><th>Товар</th><th>Реестр ПП 719</th></tr>")
for r in rows:
    m0 = marg(r, 0) if r["up"] else None; m1 = marg(r, .1) if r["up"] else None
    cls = "bad" if m1 is None or m1 < 0 else "ok" if m1 >= .15 else "warn"
    regc = (e(r["reg"])[:60] if r["reg"] else ("нет записи" if r["n"] in (4, 5, 9, 10) else "не требуется"))
    h.append(f"<tr class='{cls}'><td>{r['n']}</td><td>{e(r['name'])}</td><td class=n>{f2(r['nmc_unit'])}</td><td class=n>{r['nmc_unit']/NMC*100:.1f} %</td><td class=n>{f2(r['up'])}</td>"
             f"<td class=n>{'—' if m0 is None else f'{m0*100:.0f} %'}</td><td class=n>{'—' if m1 is None else f'{m1*100:.0f} %'}</td><td class=n>{f2(tgt(r))}</td>"
             f"<td class=small>{e((r['prod'] or '')[:85])}</td><td class=small>{regc}</td></tr>")
NP = sum(r["nmc_unit"] for r in pri); CP = sum(r["up"] for r in pri)
h.append(f"<tr><th></th><th>Найдено {len(pri)} из {len(rows)}</th><th class=n>{f2(NP)}</th><th class=n>{NP/NMC*100:.0f} %</th><th class=n>{f2(CP)}</th><th class=n>{(1-CP*(1+DELIV)/NP)*100:.0f} %</th><th class=n>{(1-CP*(1+DELIV)/(.9*NP))*100:.0f} %</th><th></th><th colspan=2>Без цены: поз. {', '.join(str(r['n']) for r in miss)} = {f2(sum(r['nmc_unit'] for r in miss))} ₽ ({sum(r['nmc_unit'] for r in miss)/NMC*100:.0f} % суммы)</th></tr></table>")
h.append(f"<div class='small'>Маржа = 1 − закупка×(1+{DELIV*100:.0f} % доставка мелкими партиями по СПб) / получаемая цена. Предел закупки = 0,85 × 0,9 × НМЦ ед. / {1+DELIV:.2f}. Объём по позициям заранее неизвестен — итоговая маржа на {f(MAXC)} ₽ зависит от того, что будут заказывать; дорогие поз. 5, 9, 10 почти наверняка составят основную часть выборки.</div>")
h.append("<!--VERDICT-->")
h.append("<h2>3. Кандидаты и почему не проходят (поз. 5, 9, 10, 11)</h2><table><tr><th>№</th><th>Что проверено</th></tr>")
for r in miss: h.append(f"<tr><td>{r['n']}</td><td class=small>{e((r['match'] + ' ' + r['note'])[:520])}</td></tr>")
h.append("</table><h2>4. Кому писать КП</h2><table><tr><th>Компания</th><th>Телефон</th><th>E-mail</th><th>Что закрывает</th><th>2 источника</th></tr>")
seen = set()
for s in SUP:
    if s.get("name") in seen: continue
    seen.add(s.get("name"))
    h.append(f"<tr><td>{e(s.get('name'))[:90]}</td><td>{e(s.get('phone'))}</td><td>{e(s.get('email'))}</td><td class=small>{e(str(s.get('closes') or ''))[:80]}</td><td>{'да' if s.get('verified') else 'нет'}</td></tr>")
h.append("</table><h2>5. Риски</h2><ul>" + "".join(f"<li>{e(r)[:240]}</li>" for r in RISK[:7]) + "</ul></body></html>")
if miss:
    VERD = (f"<b>Вывод: пока НЕ ИДЁМ — нет подтверждённого товара на {sum(r['nmc_unit'] for r in miss)/NMC*100:.0f} % суммы.</b> На поз. "
            f"{', '.join(str(r['n']) for r in miss)} не найдено модели, проходящей все показатели ТТ (а для 5, 9, 10 — ещё и с записью ПП 719). "
            f"По найденным мелким позициям маржа при −10 % — {(1-CP*(1+DELIV)/(.9*NP))*100:.0f} % (СКЛ 24 В дороже НМЦ). У ламп СКЛ (поз. 4) записи в реестре нет ни у одного производителя. "
            "Идти можно, только если до 10.11 производители (Пром-Свет/PromLED СПб, Varton, Световые технологии, Ферекс, Фотон+) подтвердят модель под ТТ с номером записи и цену не выше «предела закупки» из таблицы. "
            "Плюс запрос разъяснений: поз. 3 (жёсткие/гибкие выводы), поз. 7 (что такое «тип 2», рамка).")
else:
    VERD = "<b>Вывод:</b> см. таблицу."
S = "".join(h).replace("<!--VERDICT-->", f"<div class='box'>{VERD}</div>")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/metro.html"
open(hp, "w").write(S); html_to_pdf(hp, "../METRO_SPB_32616525105.pdf")

wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Позиции"
ws.append(["№", "Наименование", "ОКПД2", "Нацрежим", "НМЦ ед. с НДС", "Доля в сумме", "Рынок ед.", "НДС в цене", "Маржа по НМЦ", "Маржа при −10 %", "Предел закупки (15 % при −10 %)",
           "Товар (модель, производитель)", "Соответствие ТТ", "Реестр ПП 719", "Источники (продавец: цена — ссылка)", "Примечание"])
for r in rows:
    src = "\n".join(f"{s.get('seller','')}: {s.get('price')} — {s.get('url')}" for s in r["src"][:4] if isinstance(s, dict))
    ws.append([r["n"], r["name"], r["okpd"], "ограничение" if r["constr"] == "true" else "—", r["nmc_unit"], r["nmc_unit"] / NMC, r["up"], r["vat"],
               marg(r, 0) if r["up"] else None, marg(r, .1) if r["up"] else None, tgt(r), r["prod"], r["match"], r["reg"], src, r["note"]])
    X = ws.max_row
    for j in (5, 7, 11): ws.cell(X, j).number_format = "#,##0.00"
    for j in (6, 9, 10): ws.cell(X, j).number_format = "0%"
ws.append([None, "ИТОГО (сумма цен единиц)", None, None, NMC, 1, CP, None, None, None, None, f"цены найдены по {len(pri)} из {len(rows)}"])
for c in ws[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
for i, w in enumerate([4, 34, 13, 12, 12, 9, 11, 22, 10, 10, 13, 40, 60, 40, 60, 60], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "C2"
s3 = wb.create_sheet("Поставщики"); s3.append(["Компания", "ИНН", "Город", "Сайт", "Телефон", "E-mail", "Что закрывает", "2 источника"])
for s in SUP: s3.append([s.get("name"), s.get("inn"), s.get("city"), s.get("site"), s.get("phone"), s.get("email"), str(s.get("closes") or ""), "да" if s.get("verified") else "нет"])
s4 = wb.create_sheet("Риски"); s4.append(["Риск"]); [s4.append([r[:900]]) for r in RISK]; s4.column_dimensions["A"].width = 140
tidy(wb, row_h=30); wb.save("../METRO_SPB_32616525105_POZICII.xlsx")
print(len(pri), round(CP, 2), round(NP, 2), [r["n"] for r in miss])
