"""ПАО «Россети Урал» — Пермэнерго, ЗП МСП № 43712 «Поставка приборов», НМЦ 2 349 163,68 с НДС, 61 поз. (30 товаров).
Данные: deep/rprib_items.json (ТЗ, НМЦ), deep/rprib_out.json (цены, агент 04.11.2026). → ../ROSSETI_URAL_PRIBORY_43712.pdf + _POZICII.xlsx"""
import json, html, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from pdfutil import html_to_pdf
from tidy import tidy

D = "data/zakupki_0211/tz/deep/"
IT = {x["n"]: x for x in json.load(open(D + "rprib_items.json"))}
O = json.load(open(D + "rprib_out.json"))
import os
for fn in ("rprib_out2.json", "rprib_out3.json"):   # добор: заменяем товар целиком, если в доборе есть цена
    if not os.path.exists(D + fn): continue
    d2 = json.load(open(D + fn))
    for q in d2.get("products", []):
        if not isinstance(q.get("unit_price"), (int, float)): continue
        qs = set(q.get("positions") or [])
        hit = [p for p in O["products"] if set(p["positions"]) & qs]
        if hit:
            for p in hit:
                keep = sorted(set(p["positions"]) - qs)
                if keep: O["products"].append(dict(p, positions=keep))
                O["products"].remove(p)
        O["products"].append(dict(q, positions=sorted(qs), product=(q.get("product") or "") + (f" [{q['price_type']}]" if q.get("price_type") else "")))
    O["risks"] += d2.get("risks", []); O["suppliers"] += d2.get("suppliers", [])
NMC = sum(x["nmc_sum"] for x in IT.values()); NMC_NV = sum(x["sum_nv"] for x in IT.values())
BAN = {7}; LIM = set([1, 2, 3, 4, 5, 6, 8, 9, 10, 11, 14, 15, 19, 20, 21, 22, 24, 25, 26, 27, 28, 29, 31, 33, 34] + list(range(38, 47)) + list(range(53, 60)))
f = lambda v: f"{v:,.0f}".replace(",", " ") if isinstance(v, (int, float)) else "—"
e = lambda s: html.escape(str(s if s is not None else "—"))
P = []
for p in O["products"]:
    pos = [n for n in p["positions"] if n in IT]
    qty = sum(IT[n]["qty"] for n in pos); nmc = sum(IT[n]["nmc_sum"] for n in pos)
    up = p.get("unit_price") if isinstance(p.get("unit_price"), (int, float)) else None
    reg = "запрет" if set(pos) & BAN else "ограничение" if set(pos) & LIM else "—"
    P.append(dict(p, pos=pos, qty=qty, nmc=nmc, nmc_u=IT[pos[0]]["nmc_unit"], up=up, cost=up * qty if up else None, reg=reg))
pri = [p for p in P if p["up"]]; miss = [p for p in P if not p["up"]]
C = sum(p["cost"] for p in pri); NP = sum(p["nmc"] for p in pri); NM = sum(p["nmc"] for p in miss)
DELIV = 25000; DAYS = 45
def net(Pr, C_):
    pre = Pr - C_ - DELIV - 0.008 * Pr - C_ * 0.2 * DAYS / 365
    return pre - 0.15 * max(pre, 0)
# сценарий: ненайденные — по НМЦ×k (k = рынок/НМЦ найденных) и пессимизм — по НМЦ
k = C / NP; Cb = C + NM * k; Cp = C + NM

css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:8.4px;color:#111;margin:0}h1{font-size:14px;margin:0 0 3px}
h2{font-size:11px;margin:6px 0 3px;border-bottom:2px solid #305496;padding-bottom:1px}table{border-collapse:collapse;width:100%;margin:2px 0 4px}
th,td{border:1px solid #bbb;padding:1px 3px;vertical-align:top}th{background:#305496;color:#fff;text-align:left}td.n{text-align:right;white-space:nowrap}
.ok{background:#c6efce}.bad{background:#f8cbad}.warn{background:#fff2cc}.small{font-size:7.3px;color:#444}.box{border:1px solid #305496;padding:4px 6px;background:#f3f6fb;margin:4px 0}"""
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body>",
     "<h1>ПАО «Россети Урал» — Пермэнерго: поставка приборов, ЗП МСП № 43712 (ID 63091)</h1>",
     "<div class='small'>ООО «ГИДРОПАС» (УСН). Документация (новая ред.), ТЗ, опросные листы, расчёт НМЦ. Цены — открытые источники на 04.11.2026 (дилеры/розница), КП не запрашивались; реестр ПП 719 — выгрузка ГИСП 01.11.2026. Все ссылки — в Excel.</div><h2>1. Условия</h2><table>"]
for a, b in [("НМЦ", f"{f(NMC)} ₽ с НДС / {f(NMC_NV)} ₽ без НДС; 61 позиция = 30 разных товаров (дубли по кварталам); частичная подача не допускается"),
             ("Подача", "corp.roseltorg.ru, до 10.11.2026 09:00 МСК; разъяснения до 09.11 15:00; только МСП; обеспечения заявки и договора нет"),
             ("Оценка", "цена 70 %, опыт поставок за 36 мес. 15 %, полномочия (письма) производителей 15 %"),
             ("Поставка", "Пермский край, авто силами поставщика: I кв. — 45 дн. с договора (19 поз.); II кв. — 16–27.02.2027 (19); III–IV кв. — 27.04–08.05.2027 (23)"),
             ("Оплата", "7 раб. дней после приёмки"),
             ("Нацрежим", "поз. 7 (Т-1014Р) — ЗАПРЕТ иностранного; поз. 1–6, 8–11, 14–15, 19–22, 24–29, 31, 33–34, 38–46, 53–59 — ОГРАНИЧЕНИЕ (номер записи ПП 719/ЕАЭС в ценовом предложении, иначе заявка = иностранная)")]:
    h.append(f"<tr><th style='width:11%'>{a}</th><td>{e(b)}</td></tr>")
h.append("</table><h2>2. Товары (сгруппированы): НМЦ против рынка</h2><table><tr><th>Товар</th><th>Поз.</th><th>Кол.</th><th>НМЦ ед.</th><th>Рынок ед.</th><th>НМЦ сумма</th><th>Рынок сумма</th><th>Рынок/НМЦ</th><th>Нацрежим / реестр</th><th>Товар, производитель</th></tr>")
for p in sorted(P, key=lambda p: -p["nmc"]):
    r = p["cost"] / p["nmc"] if p["cost"] else None
    cls = "bad" if r is None or r > 0.9 else "warn" if r > 0.75 else "ok"
    h.append(f"<tr class='{cls}'><td>{e(p['key'])}</td><td class=small>{', '.join(map(str, p['pos']))}</td><td class=n>{p['qty']:g}</td><td class=n>{f(p['nmc_u'])}</td><td class=n>{f(p['up'])}</td><td class=n>{f(p['nmc'])}</td><td class=n>{f(p['cost'])}</td>"
             f"<td class=n>{'—' if r is None else f'{r*100:.0f} %'}</td><td class=small>{e(p['reg'])}{'; запись ' + e(p['registry']) if p.get('registry') else ('; записи нет' if p['reg'] != '—' else '')}</td><td class=small>{e(str(p.get('product'))[:80])}</td></tr>")
h.append(f"<tr><th colspan=5>Найдено {len(pri)} из {len(P)} товаров</th><th class=n>{f(NP)}</th><th class=n>{f(C)}</th><th class=n>{C/NP*100:.0f} %</th><th colspan=2>Без цены: {f(NM)} ₽ ({NM/NMC*100:.0f} % НМЦ)</th></tr></table>")
h.append("<h2>3. Маржа (наша цена = без НДС; закупка с НДС; доставка ≈25 тыс.)</h2><table><tr><th>Сценарий</th><th>Себестоимость</th><th>Маржа по НМЦ б/НДС</th><th>При −10 %</th><th>Чистая по НМЦ</th><th>Чистая при −10 %</th></tr>")
for nm, Cx in [(f"База: ненайденные по тому же соотношению рынок/НМЦ ({k*100:.0f} %)", Cb), ("Пессимизм: ненайденные по НМЦ с НДС", Cp)]:
    m1 = (NMC_NV * .9 - Cx - DELIV) / (NMC_NV * .9)
    h.append(f"<tr class='{'ok' if m1 >= .15 else 'warn' if m1 > 0 else 'bad'}'><td>{e(nm)}</td><td class=n>{f(Cx + DELIV)}</td><td class=n>{(NMC_NV - Cx - DELIV) / NMC_NV * 100:.0f} %</td><td class=n>{m1*100:.0f} %</td><td class=n>{f(net(NMC_NV, Cx))}</td><td class=n>{f(net(NMC_NV * .9, Cx))}</td></tr>")
h.append("</table>")
ikz = next(p for p in P if p["pos"] == [35])
lim_ikz = (NMC_NV * 0.9 * 0.85 - (Cb - ikz["nmc"] * k) - DELIV)  # сколько можно отдать за 2 комплекта ИКЗ, чтобы маржа при −10 % была 15 %
h.append(f"<div class='box'><b>Вывод: НЕ ИДЁМ по открытым ценам; решение — только после КП «Антракс» по ИКЗ-В43Л.</b> Рынок найденных товаров = {C/NP*100:.0f} % их НМЦ (с НДС) — это выше НМЦ без НДС, т.е. по ним маржи нет "
         f"(розница; бинокли дороже НМЦ). Всё решают 2 комплекта ИКЗ-В43Л (39 % НМЦ, только КП) и колонки «Хроматэк». Для 15 % при −10 % за 2 комплекта ИКЗ можно заплатить не более ≈{f(lim_ikz)} ₽ с НДС (НМЦ {f(ikz['nmc'])}). "
         "Нацрежим: поз. 7 Т-1014Р — ЗАПРЕТ, записи в реестре нет → поставить нельзя (заявку отклонят). По ограничению записей ПП 719 почти ни у кого нет (есть у ИКЗ-В43Л 10457518, но в исполнении У3, а ОЛ требует УХЛ1); "
         "при российском конкуренте с записями наша заявка проигрывает. Плюс 15 % оценки — письма производителей, у нас их нет.</div>")
h.append("<h2>4. Кому писать КП</h2><table><tr><th>Компания</th><th>Телефон</th><th>E-mail</th><th>Что</th></tr>")
for s in O["suppliers"][:9]:
    h.append(f"<tr><td>{e(s.get('name'))[:80]}</td><td>{e(s.get('phone'))[:40]}</td><td>{e(s.get('email'))}</td><td class=small>{e(s.get('closes'))[:60]}</td></tr>")
h.append("</table><h2>5. Риски</h2><ul>" + "".join(f"<li>{e(r if isinstance(r, str) else json.dumps(r, ensure_ascii=False))[:230]}</li>" for r in O["risks"][:7]) + "</ul></body></html>")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/rprib.html"
open(hp, "w").write("".join(h)); html_to_pdf(hp, "../ROSSETI_URAL_PRIBORY_43712.pdf")

wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Позиции (61)"
ws.append(["№", "Номенкл.", "Наименование", "Ед.", "Кол-во", "Срок", "НМЦ ед. с НДС", "НМЦ сумма", "Нацрежим", "Товар (рынок)", "Рынок ед.", "Рынок сумма", "НДС", "Реестр ПП 719", "Соответствие", "Источники", "Примечание"])
byn = {n: p for p in P for n in p["pos"]}
for n in sorted(IT):
    x = IT[n]; p = byn.get(n, {})
    src = "\n".join(f"{s.get('seller','')}: {s.get('price')} — {s.get('url')}" for s in (p.get("sources") or [])[:3] if isinstance(s, dict))
    ws.append([n, x["nom"], x["name"], x["unit"], x["qty"], x["date"], x["nmc_unit"], x["nmc_sum"], "запрет" if n in BAN else "ограничение" if n in LIM else "—", p.get("product"),
               p.get("up"), (p.get("up") or 0) * x["qty"] or None, p.get("vat"), p.get("registry"), p.get("match"), src, p.get("note")])
    for j in (7, 8, 11, 12): ws.cell(ws.max_row, j).number_format = "#,##0.00"
for c in ws[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
for i, w in enumerate([4, 9, 40, 6, 6, 18, 11, 12, 11, 40, 10, 12, 14, 16, 30, 50, 50], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "D2"; ws.auto_filter.ref = f"A1:Q{ws.max_row}"
s3 = wb.create_sheet("Поставщики"); s3.append(["Компания", "ИНН", "Город", "Сайт", "Телефон", "E-mail", "Позиции", "Подтверждение"])
for s in O["suppliers"]: s3.append([s.get("name"), s.get("inn"), s.get("city"), s.get("site"), s.get("phone"), s.get("email"), str(s.get("closes")), str(s.get("verified"))])
s4 = wb.create_sheet("Риски"); s4.append(["Риск"]); [s4.append([(r if isinstance(r, str) else json.dumps(r, ensure_ascii=False))[:900]]) for r in O["risks"]]; s4.column_dimensions["A"].width = 140
tidy(wb, row_h=26); wb.save("../ROSSETI_URAL_PRIBORY_43712_POZICII.xlsx")
print(round(C), round(NP), round(NM), round(k, 3), round(Cb), round(Cp), round(lim_ikz))
