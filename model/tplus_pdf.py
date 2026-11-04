"""ПАО «Т Плюс» (Кировский филиал) 32616484848 / B2B № 4644566 — электротехническая продукция, 182 поз.
Данные: data/zakupki_0211/tz/deep/tplus_items.json (ТЗ после изменений 02.11.2026) + tplus_g{1,2,3}_out.json (цены, агенты 04.11.2026).
→ ../TPLUS_KIROV_32616484848.pdf (2 стр.) + ../TPLUS_KIROV_32616484848_POZICII.xlsx"""
import json, html, os, collections, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from pdfutil import html_to_pdf
from tidy import tidy

D = "data/zakupki_0211/tz/deep/"
IT = json.load(open(D + "tplus_items.json"))
PR, MK, RISK = {}, [], []
for g in (1, 2, 3):
    p = D + f"tplus_g{g}_out.json"
    if not os.path.exists(p): continue
    o = json.load(open(p))
    for x in o.get("items", []): PR[int(x["n"])] = x
    MK += o.get("makers", []); RISK += [r if isinstance(r, str) else json.dumps(r, ensure_ascii=False) for r in o.get("risks", [])]
NMC = sum(x["nmc_sum"] for x in IT)                 # 1 265 686 с НДС
NMC_NV = sum(x["plan_nv"] * x["qty"] for x in IT)   # 1 037 448 без НДС — с ней сравнивают цену (оценка без НДС)
DELIV = 30000   # оценка: ~10 отгрузок ТК в Киров/Кирово-Чепецк (даты поставки апр.–ноя. 2027) по ~3 тыс.
DAYS = 45       # закупка перед поставкой; оплата 30 дн. (102 поз.) или 7 раб. дн. (74 поз.) после поставки

def net(P, C, days=DAYS):
    pre = P - C - DELIV - 0.008 * P - C * 0.2 * days / 365
    return pre - 0.15 * max(pre, 0)

rows = []
for x in IT:
    p = PR.get(x["n"], {}); up = p.get("unit_price")
    try: up = float(up) if up not in (None, "") else None
    except (TypeError, ValueError): up = None
    rows.append(dict(x, up=up, prod=p.get("product") or "", match=p.get("match") or "", vat=p.get("vat") or "",
                     src=p.get("sources") or [], kp=p.get("supplier_for_kp") or "", note=p.get("note") or ""))
pri = [r for r in rows if r["up"]]
miss = [r for r in rows if not r["up"]]
C_pr = sum(r["up"] * r["qty"] for r in pri)
NMC_pr = sum(r["nmc_sum"] for r in pri); NV_pr = sum(r["plan_nv"] * r["qty"] for r in pri)
NV_miss = sum(r["plan_nv"] * r["qty"] for r in miss); NMC_miss = sum(r["nmc_sum"] for r in miss)
k_pr = C_pr / NV_pr if NV_pr else 0            # рынок / НМЦ без НДС по найденным
C_base = C_pr + NV_miss * k_pr                  # ненайденные — тем же соотношением, что найденные
C_pess = C_pr + NMC_miss                        # ненайденные — по НМЦ с НДС (маржа 0 и хуже)

# ─── PDF ───
f = lambda v: f"{v:,.0f}".replace(",", " ")
e = lambda s: html.escape(str(s if s is not None else "—"))
css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:8.4px;color:#111;margin:0}h1{font-size:14px;margin:0 0 3px}
h2{font-size:11px;margin:6px 0 3px;border-bottom:2px solid #305496;padding-bottom:1px}table{border-collapse:collapse;width:100%;margin:2px 0 4px}
th,td{border:1px solid #bbb;padding:1px 3px;vertical-align:top}th{background:#305496;color:#fff;text-align:left}td.n{text-align:right;white-space:nowrap}
.ok{background:#c6efce}.bad{background:#f8cbad}.warn{background:#fff2cc}.small{font-size:7.4px;color:#444}.box{border:1px solid #305496;padding:4px 6px;background:#f3f6fb;margin:4px 0}"""
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body>",
     "<h1>ПАО «Т Плюс», Кировский филиал — электротехническая продукция, ЕИС 32616484848 (B2B-Т Плюс № 4644566)</h1>",
     "<div class='small'>ООО «ГИДРОПАС» (УСН). Документация + изменения от 02.11.2026 (протокол ЗК № 50300-18-08/1567). Цены — открытые источники на 04.11.2026, с НДС; КП не запрашивались. Все позиции и ссылки — в Excel.</div>",
     "<h2>1. Условия и что изменилось</h2><table>"]
cond = [("НМЦ", f"{f(NMC)} ₽ с НДС / {f(NMC_NV)} ₽ без НДС; 182 позиции, 2 454 ед., один лот, весь объём"),
        ("Изменения 02.11", "Исключена поз. 4 «Блок питания ВЧ поста ПВ3-90М1» (2 шт., 135 672 ₽ с НДС); НМЦ снижена с 1 401 358 до 1 265 686 ₽; позиций 183 → 182"),
        ("Подача", "до 10.11.2026 15:00 МСК, ЭТП B2B-Т Плюс (tender.tplusgroup.ru); вскрытие 11.11 10:00; итоги — 16.12.2026; переторжка (несколько ценовых шагов, +30 мин)"),
        ("Оценка", "цена 80 % (сравнивается БЕЗ НДС, дисконтируется на отсрочку по ключевой ставке ЦБ), условия оплаты 20 %; заявка в Excel/Word + PDF"),
        ("Обеспечение", "заявки и договора — нет"),
        ("Поставка", "авто, 4 адреса: Киров (Ломоносова 2а — 68 поз., Луганская 51а — 57, пр. Колесникова 6 — 40), Кирово-Чепецк (Рабочий пер. 4 — 17); даты по ТЗ: апр. 2027 — 28 поз., май — 106, июнь — 22, июль–нояб. — 26"),
        ("Оплата", "102 поз. — 100 % с отсрочкой 30 дней; 74 поз. — 100 % в течение 7 раб. дней после поставки; аванса нет"),
        ("Договор", "протокол разногласий НЕ допускается (только альтернативные условия оплаты); пени 0,1 %/день; расторжение при просрочке > 10 дней; смена налогового режима цену не меняет"),
        ("Нацрежим", "«ограничение» (ПП 1875) по 11 позициям — нужен номер реестровой записи ПП 719 или ЕАЭС; по остальным — нет")]
for k, v in cond: h.append(f"<tr><th style='width:13%'>{k}</th><td>{e(v)}</td></tr>")
h.append("</table>")

# группы
def grp(n):
    s = n.upper()
    for k, g in [("ВЫКЛЮЧАТ", "Выключатели, переключатели"), ("ПЕРЕКЛЮЧ", "Выключатели, переключатели"), ("РОЗЕТ", "Розетки, вилки, удлинители"), ("ВИЛК", "Розетки, вилки, удлинители"),
                 ("УДЛИН", "Розетки, вилки, удлинители"), ("СОЕДИНИТЕЛЬ ШТЕП", "Розетки, вилки, удлинители"), ("ПРЕДОХР", "Предохранители, вставки"), ("ВСТАВК", "Предохранители, вставки"),
                 ("ЗАЖИМ", "Зажимы, клеммы, колодки"), ("КЛЕММ", "Зажимы, клеммы, колодки"), ("КОЛОДК", "Зажимы, клеммы, колодки"), ("КОРОБК", "Коробки, боксы, сальники"),
                 ("БОКС", "Коробки, боксы, сальники"), ("САЛЬНИК", "Коробки, боксы, сальники"), ("ПРУЖИН", "Запчасти разъединителей/трансформаторов/генераторов"),
                 ("ЛАМЕЛ", "Запчасти разъединителей/трансформаторов/генераторов"), ("РЕМКОМПЛ", "Запчасти разъединителей/трансформаторов/генераторов"),
                 ("ПРОКЛАДК", "Запчасти разъединителей/трансформаторов/генераторов"), ("КРЫШК", "Запчасти разъединителей/трансформаторов/генераторов"),
                 ("ДАТЧИК", "Датчики, блок-контакты, посты"), ("БЛОК-КОНТ", "Датчики, блок-контакты, посты"), ("ПОСТ", "Датчики, блок-контакты, посты"),
                 ("СТАРТЕР", "Светотехника, ЭПРА, стартеры"), ("ЭПРА", "Светотехника, ЭПРА, стартеры"), ("ПРА", "Светотехника, ЭПРА, стартеры"), ("ЛАМП", "Светотехника, ЭПРА, стартеры"), ("ПАТРОН", "Светотехника, ЭПРА, стартеры")]:
        if s.lstrip("~").startswith(k) or (k in s[:40]): return g
    return "Прочее"
G = collections.OrderedDict()
for r in rows:
    r["grp"] = grp(r["name"]); g = G.setdefault(r["grp"], dict(n=0, np=0, nmc=0, nv=0, nvp=0, c=0))
    g["n"] += 1; g["nmc"] += r["nmc_sum"]; g["nv"] += r["plan_nv"] * r["qty"]
    if r["up"]: g["np"] += 1; g["nvp"] += r["plan_nv"] * r["qty"]; g["c"] += r["up"] * r["qty"]
h.append("<h2>2. Себестоимость по группам (рынок с НДС против НМЦ без НДС по найденным позициям)</h2><table><tr><th>Группа</th><th>Поз.</th><th>Найдено цен</th><th>НМЦ с НДС</th><th>НМЦ без НДС</th><th>НМЦ без НДС найденных</th><th>Рынок найденных</th><th>Рынок / НМЦ б/НДС</th></tr>")
for k, g in sorted(G.items(), key=lambda t: -t[1]["nmc"]):
    ratio = g["c"] / g["nvp"] if g["nvp"] else None
    cls = "" if ratio is None else "ok" if ratio <= 0.7 else "warn" if ratio <= 0.9 else "bad"
    h.append(f"<tr class='{cls}'><td>{e(k)}</td><td class=n>{g['n']}</td><td class=n>{g['np']}</td><td class=n>{f(g['nmc'])}</td><td class=n>{f(g['nv'])}</td><td class=n>{f(g['nvp'])}</td><td class=n>{f(g['c'])}</td><td class=n>{'—' if ratio is None else f'{ratio*100:.0f} %'}</td></tr>")
h.append(f"<tr><th>Итого</th><th class=n>{len(rows)}</th><th class=n>{len(pri)}</th><th class=n>{f(NMC)}</th><th class=n>{f(NMC_NV)}</th><th class=n>{f(NV_pr)}</th><th class=n>{f(C_pr)}</th><th class=n>{k_pr*100:.0f} %</th></tr></table>")

h.append("<h2>3. Маржа и чистая прибыль (наша цена = цена без НДС, мы на УСН)</h2><table><tr><th>Сценарий себестоимости</th><th>Себестоимость</th><th>Маржа по НМЦ б/НДС</th><th>Маржа при −10 %</th><th>Маржа при −20 %</th><th>Чистая по НМЦ</th><th>Чистая при −10 %</th><th>Чистая при −20 %</th></tr>")
for name, C in [(f"База: найдено {len(pri)} поз. по рынку, остальные {len(miss)} — тем же % от НМЦ", C_base),
                (f"Пессимизм: ненайденные {len(miss)} поз. — по НМЦ с НДС", C_pess)]:
    P0, P1, P2 = NMC_NV, NMC_NV * 0.9, NMC_NV * 0.8; Ct = C + DELIV
    m1 = (P1 - Ct) / P1; cls = "ok" if m1 >= 0.15 else "warn" if m1 > 0 else "bad"
    h.append(f"<tr class='{cls}'><td>{e(name)}</td><td class=n>{f(Ct)}</td>" + "".join(f"<td class=n>{(P - Ct) / P * 100:.1f} %</td>" for P in (P0, P1, P2)) +
             "".join(f"<td class=n>{f(net(P, C))}</td>" for P in (P0, P1, P2)) + "</tr>")
h.append(f"</table><div class='small'>Себестоимость включает доставку ≈{f(DELIV)} ₽ (оценка: ~10 отгрузок ТК). Чистая: цена − закупка − доставка − 0,8 % ЭТП − деньги 20 % год. × {DAYS} дн. − УСН 15 %. Сравнение с НМЦ БЕЗ НДС, т.к. Т Плюс оценивает цены без НДС.</div>")
h.append("<!--VERDICT-->")

# крупнейшие позиции
top = sorted(rows, key=lambda r: -r["nmc_sum"])[:14]
h.append("<h2>4. Крупнейшие позиции (≈ половина НМЦ)</h2><table><tr><th>№</th><th>Позиция</th><th>Кол.</th><th>НМЦ ед. б/НДС</th><th>Рынок ед. с НДС</th><th>Товар / источник</th></tr>")
for r in top:
    h.append(f"<tr><td>{r['n']}</td><td>{e(' '.join(r['name'].split())[:70])}</td><td class=n>{r['qty']}</td><td class=n>{f(r['plan_nv'])}</td><td class=n>{f(r['up']) if r['up'] else '—'}</td><td class=small>{e((r['prod'] or r['kp'])[:90])}</td></tr>")
h.append("</table>")
h.append("<h2>5. Кому писать КП</h2><table><tr><th>Компания</th><th>Город</th><th>Телефон</th><th>E-mail</th><th>Сайт</th><th>Что закрывает</th></tr>")
seen = set()
for m in MK:
    if m.get("name") in seen: continue
    seen.add(m.get("name"))
    if len(seen) > 9: break
    h.append(f"<tr><td>{e(m.get('name'))}</td><td>{e(m.get('city'))}</td><td>{e(str(m.get('phone') or '').split(';')[0])}</td><td>{e(m.get('email'))}</td><td>{e(m.get('site'))}</td><td class=small>{e(str(m.get('closes') or '')[:90])}</td></tr>")
h.append("</table><h2>6. Риски</h2><ul>" + "".join(f"<li>{e(r)[:230]}</li>" for r in RISK[:8]) + "</ul></body></html>")

m_base10 = (NMC_NV * 0.9 - C_base - DELIV) / (NMC_NV * 0.9)
m_pess10 = (NMC_NV * 0.9 - C_pess - DELIV) / (NMC_NV * 0.9)
VERD = os.environ.get("VERDICT") or (f"<b>Вывод:</b> база — маржа при −10 % {m_base10*100:.0f} %, пессимизм — {m_pess10*100:.0f} %.")
S = "".join(h).replace("<!--VERDICT-->", f"<div class='box'>{VERD}</div>")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/tplus.html"
open(hp, "w").write(S); html_to_pdf(hp, "../TPLUS_KIROV_32616484848.pdf")

# ─── Excel ───
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Позиции (182)"
ws.append(["№", "Группа", "Код SAP", "Позиция ТЗ", "Ед.", "Кол-во", "НМЦ ед. с НДС", "НМЦ ед. без НДС", "НМЦ сумма с НДС", "Рынок ед. (с НДС)", "НДС в цене", "Рынок сумма",
           "Рынок / НМЦ б/НДС", "Товар (модель, производитель)", "Соответствие ТЗ", "Дата поставки", "Адрес", "Нацрежим", "Источники (цена — ссылка)", "Кому запросить КП", "Примечание"])
for r in rows:
    src = "\n".join(f"{s.get('seller','')}: {s.get('price')} — {s.get('url')}" for s in r["src"][:3] if isinstance(s, dict))
    ws.append([r["n"], r["grp"], r.get("sap"), " ".join(r["name"].split()), r["unit"], r["qty"], r["nmc_unit"], r["plan_nv"], r["nmc_sum"], r["up"], r["vat"],
               (r["up"] or 0) * r["qty"] or None, (r["up"] / r["plan_nv"]) if r["up"] and r["plan_nv"] else None, r["prod"], r["match"], r["date"], r.get("place"), r.get("restr"), src, r["kp"], r["note"]])
    X = ws.max_row
    for j in (7, 8, 9, 10, 12): ws.cell(X, j).number_format = "#,##0.00"
    ws.cell(X, 13).number_format = "0%"
for c in ws[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
for i, w in enumerate([5, 20, 9, 55, 6, 7, 11, 11, 12, 11, 10, 12, 9, 34, 16, 11, 26, 11, 50, 30, 30], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "E2"; ws.auto_filter.ref = f"A1:U{ws.max_row}"
s2 = wb.create_sheet("Итог")
for row in [["Показатель", "Значение"], ["НМЦ с НДС", NMC], ["НМЦ без НДС (база сравнения)", NMC_NV], ["Позиций с ценой", len(pri)], ["Позиций без цены", len(miss)],
            ["Рынок по найденным, с НДС", C_pr], ["НМЦ без НДС по найденным", NV_pr], ["Себестоимость база (+доставка)", C_base + DELIV], ["Себестоимость пессимизм (+доставка)", C_pess + DELIV],
            ["Маржа при −10 %, база", m_base10], ["Маржа при −10 %, пессимизм", m_pess10], ["Чистая при −10 %, база", net(NMC_NV * 0.9, C_base)], ["Чистая при −10 %, пессимизм", net(NMC_NV * 0.9, C_pess)]]:
    s2.append(row)
for c in s2["B"][1:]: c.number_format = "0.0%" if isinstance(c.value, float) and abs(c.value) < 1.5 else "#,##0"
s2.column_dimensions["A"].width = 40; s2.column_dimensions["B"].width = 18
s3 = wb.create_sheet("Поставщики для КП")
s3.append(["Компания", "ИНН", "Город", "Сайт", "Телефон", "E-mail", "Что закрывает", "Подтверждено в 2 источниках"])
for m in MK: s3.append([m.get("name"), m.get("inn"), m.get("city"), m.get("site"), m.get("phone"), m.get("email"), str(m.get("closes") or "")[:300], "да" if m.get("verified") else "нет"])
s4 = wb.create_sheet("Риски"); s4.append(["Риск"]); [s4.append([r[:800]]) for r in RISK]; s4.column_dimensions["A"].width = 140
tidy(wb, row_h=24); wb.save("../TPLUS_KIROV_32616484848_POZICII.xlsx")
print(dict(priced=len(pri), miss=len(miss), C_pr=round(C_pr), NV_pr=round(NV_pr), k=round(k_pr, 3), C_base=round(C_base), C_pess=round(C_pess), m10b=round(m_base10, 3), m10p=round(m_pess10, 3),
           net10b=round(net(NMC_NV * .9, C_base)), net10p=round(net(NMC_NV * .9, C_pess))))
