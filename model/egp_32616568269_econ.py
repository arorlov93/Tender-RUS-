"""Экономика закупки ЕИС 32616568269, ПАО «ЕГП» (Ессентуки): панели Ecophon Соло PE Круг, светильники и шинопровод Geniled.
Цены: data/tz_lots/egp_32616568269/market_prices.json (страницы открыты 11.11.2026). → ../EGP_32616568269_EKONOMIKA.xlsx + .pdf"""
import json, html, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from isp_letterhead import SCRATCH
import subprocess

D = "data/tz_lots/egp_32616568269/market_prices.json"
P = {p["pos"]: p for p in json.load(open(D))["positions"]}
NMC = 4_522_804.97
# допущения (не рыночные данные, а оценки для проверки)
DELIVERY = 120_000      # оценка: 1 тент 82 м³ Москва → Ессентуки, объём ~30 м³; запросить ТК
BG_COST = 15_000        # оценка: гарантия на возврат аванса 30%, ~3 мес, ~3-4% годовых, мин. комиссия
OTHER = 20_000          # оценка: упаковка/страховка/образец, курьер документов

NAMES = {1: "Панель Соло PE Круг D1180, цвет по RAL", 2: "Панель Соло PE Круг D1180, RAL 9003", 3: "Панель круг D600 (Соло PE D600 не выпускается)",
         4: "Комплект подвеса (трос 2 м, фиксатор, карабин, спираль)", 5: "Geniled Art Ring RE d1600 120 Вт 4000К Ra90", 6: "Geniled Track Classic Zoom RE 20 Вт 4000К Ra90",
         7: "Шинопровод однофазный Geniled 2000 черный", 8: "Ввод питания Geniled черный", 9: "Заглушка Geniled черная", 10: "Подвес Geniled 2 м для шинопровода"}

def s(p, k): return P[p]["stats"][k]
# сценарии цены закупки за шт
SC = {
 "A. Точные модели, мин. цена": {1: s(1, "exact_min"), 2: s(2, "exact_min"), 3: 6710.0, 4: s(4, "exact_min"), 5: s(5, "exact_min"), 6: s(6, "exact_min"), 7: s(7, "exact_min"), 8: s(8, "exact_min"), 9: s(9, "exact_min"), 10: s(10, "exact_min")},
 "B. Точные модели, медиана": {1: s(1, "exact_median"), 2: s(2, "exact_median"), 3: 8186.0, 4: s(4, "exact_median"), 5: s(5, "exact_median"), 6: s(6, "exact_median"), 7: s(7, "exact_median"), 8: s(8, "exact_median"), 9: s(9, "exact_median"), 10: s(10, "exact_median")},
 "C. С аналогами, мин. цена": {1: round(6820 * 1.333), 2: s(2, "all_min"), 3: s(3, "all_min"), 4: s(4, "all_min"), 5: s(5, "all_min"), 6: s(6, "all_min"), 7: s(7, "all_min"), 8: s(8, "all_min"), 9: s(9, "all_min"), 10: s(10, "all_min")},
}
NOTE3 = {"A. Точные модели, мин. цена": "D800 Соло PE 6 710 как ориентир (D600 только «Соло Фридом» по запросу)",
         "B. Точные модели, медиана": "D800 Соло PE 8 186", "C. С аналогами, мин. цена": "K&R OST Круг D600, каменная вата, КМ1, 4 935"}

def calc(sc):
    rows = [(k, NAMES[k], P[k]["qty"], v, P[k]["qty"] * v) for k, v in SC[sc].items()]
    goods = sum(r[4] for r in rows)
    cost = goods + DELIVERY + BG_COST + OTHER
    return rows, goods, cost

RES = {sc: calc(sc) for sc in SC}
DROPS = [0, 0.05, 0.10, 0.15, 0.20]

# ---------- Excel ----------
wb = Workbook(); ws = wb.active; ws.title = "Экономика"
B = Font(bold=True); H = PatternFill("solid", fgColor="0B2545"); HF = Font(bold=True, color="FFFFFF")
ws.append(["Закупка ЕИС 32616568269, ПАО «ЕГП», Ессентуки. НМЦ, ₽:", NMC]); ws["A1"].font = B
ws.append(["Цены рынка на 11.11.2026, с НДС (Гидропас на УСН, НДС к вычету нет). Доставка, гарантия, прочее: оценки."]); ws.append([])
for sc, (rows, goods, cost) in RES.items():
    ws.append([sc]); ws.cell(ws.max_row, 1).font = B
    ws.append(["№", "Позиция", "Кол-во", "Цена за шт, ₽", "Сумма, ₽"])
    for c in ws[ws.max_row]: c.fill = H; c.font = HF
    for r in rows: ws.append(list(r))
    for lbl, v in [("Товар итого", goods), ("Доставка до Ессентуков (оценка)", DELIVERY), ("Банковская гарантия на аванс (оценка)", BG_COST), ("Прочее (оценка)", OTHER), ("СЕБЕСТОИМОСТЬ", cost)]:
        ws.append(["", lbl, "", "", round(v, 2)]); ws.cell(ws.max_row, 2).font = B
    ws.append(["", "Снижение от НМЦ", "Цена заявки, ₽", "Маржа до налога, ₽", "Маржа, %"])
    for c in ws[ws.max_row]: c.font = B
    for d in DROPS:
        bid = NMC * (1 - d); ws.append(["", f"−{int(d*100)}%", round(bid, 2), round(bid - cost, 2), round((bid - cost) / bid * 100, 1)])
    ws.append([])
ws2 = wb.create_sheet("Источники")
ws2.append(["№", "Продавец", "Цена, ₽", "Точная/аналог", "Наличие", "URL", "Примечание"])
for k in sorted(P):
    for o in P[k]["offers"]:
        ws2.append([k, o["seller"], o["price_rub_with_vat"], o["exact_or_analogue"], o["in_stock"], o["url"], o.get("note", "")])
for w, col in zip([6, 52, 12, 16, 16], "ABCDE"): ws.column_dimensions[col].width = w
for w, col in zip([5, 28, 12, 12, 14, 60, 60], "ABCDEFG"): ws2.column_dimensions[col].width = w
wb.save("../EGP_32616568269_EKONOMIKA.xlsx")

# ---------- PDF ----------
e = html.escape
def money(v): return f"{v:,.0f}".replace(",", " ")
a_rows, a_goods, a_cost = RES["A. Точные модели, мин. цена"]
h = ["<html><head><meta charset=utf-8><style>@page{size:A4;margin:12mm 12mm}body{font-family:'DejaVu Sans',Arial;font-size:8.6px;color:#111}h1{font-size:13px;margin:0 0 3px;color:#0B2545}h2{font-size:10px;margin:9px 0 3px;color:#0B2545}"
     "table{border-collapse:collapse;width:100%;margin-bottom:4px}th{background:#0B2545;color:#fff;padding:3px 4px;text-align:left}td{border-bottom:1px solid #ddd;padding:2.5px 4px}td.n{text-align:right;white-space:nowrap}.g{color:#555}.r{color:#B00020;font-weight:bold}.ok{color:#1B7F3B;font-weight:bold}ul{margin:2px 0 2px 14px;padding:0}li{margin:1px 0}</style></head><body>",
     f"<h1>Экономика закупки № 32616568269 · ПАО «ЕГП», Ессентуки</h1><div class=g>Акустические панели, светильники, шинопровод · НМЦ <b>{money(NMC)} ₽</b> · подача до 18.11.2026 · только МСП · цены рынка на 11.11.2026</div>"]
h.append("<h2>1. Себестоимость по позициям (сценарий A: точные модели, минимальная цена)</h2><table><tr><th>№</th><th>Позиция</th><th>Кол.</th><th>Цена, ₽</th><th>Сумма, ₽</th><th>Медиана, ₽</th><th>С аналогами, ₽</th></tr>")
for k, nm, q, v, sm in a_rows:
    h.append(f"<tr><td>{k}</td><td>{e(nm)}</td><td class=n>{q}</td><td class=n>{money(v)}</td><td class=n>{money(sm)}</td><td class=n>{money(SC['B. Точные модели, медиана'][k])}</td><td class=n>{money(SC['C. С аналогами, мин. цена'][k])}</td></tr>")
h.append("</table>")
h.append("<h2>2. Итог по трём сценариям</h2><table><tr><th>Сценарий</th><th>Товар</th><th>Себестоимость*</th>" + "".join(f"<th>−{int(d*100)}%</th>" for d in DROPS) + "</tr>")
for sc, (rows, goods, cost) in RES.items():
    cells = "".join(f"<td class=n>{money(NMC*(1-d)-cost)}<br><span class=g>{(NMC*(1-d)-cost)/(NMC*(1-d))*100:.0f}%</span></td>" for d in DROPS)
    h.append(f"<tr><td>{e(sc)}</td><td class=n>{money(goods)}</td><td class=n><b>{money(cost)}</b></td>{cells}</tr>")
h.append(f"</table><div class=g>* включая оценки: доставка {money(DELIVERY)} ₽, гарантия на аванс {money(BG_COST)} ₽, прочее {money(OTHER)} ₽. В ячейках: маржа до налога, ₽ и % от цены заявки. Налог УСН не вычтен.</div>")
h.append("<h2>3. Выводы</h2><ul>"
 f"<li><b>По розничным ценам маржи почти нет:</b> себестоимость по точным моделям {money(a_cost)} ₽ ({a_cost/NMC*100:.0f}% от НМЦ), по медиане выше НМЦ. НМЦ, судя по всему, посчитана по тем же розничным ценам. Заработать можно только на оптовых ценах дилеров и аналогах: со сценарием C маржа {money(NMC*0.9-RES['C. С аналогами, мин. цена'][2])} ₽ при снижении 10%.</li>"
 "<li><b>Решение об участии — только после оптовых КП.</b> Ориентир: себестоимость не выше 3,6 млн ₽, тогда при снижении 10–15% остаётся 8–13%.</li>"
 "<li><b>Главная статья — кольцевые светильники Ring d1600</b> (35 шт × 52–55 тыс. ₽ ≈ 1,85–1,94 млн ₽). Нужен оптовый запрос дилеру Geniled: розничная цена при 35 шт обычно снижается.</li>"
 "<li><b>Панели D600:</b> «Соло PE Круг» D600 не выпускается (только D1180 и D800). Варианты: «Соло Фридом» по чертежу (цена по запросу у Сен-Гобен/дилера) или аналог K&R OST D600, КМ1. Нужно разъяснение заказчика.</li>"
 "<li><b>Цвет RAL дороже на 33%</b> (прайс E-phon: 7 686 → 10 248 ₽).</li>"
 "<li><b>Противоречия ТЗ:</b> Ring: в названии 120 Вт, в требованиях не более 96 Вт при потоке ≥ 8 160 лм. Трековый: в названии чёрный, в характеристиках белый. Aeq ≥ 2,6 м² указан для панели 1192×1192, к кругу D600 неприменим. Всё это в запрос разъяснений.</li>"
 "</ul><h2>4. Риски договора</h2><ul>"
 "<li><span class=r>Штраф 200 000 ₽</span> за каждое несвоевременное или неверное оформление первичных документов (п. 6.3). Это больше маржи на отдельных позициях: документы только через бухгалтера, УПД без ошибок.</li>"
 "<li>Аванс 30% только после банковской гарантии на 100% аванса, банк из списка заказчика. Без гарантии закупаем на свои: ≈ 3,0–3,7 млн ₽ на 60 дней.</li>"
 "<li>Пени 1/300 ключевой ставки в день от цены всего договора. Срок поставки 60 дней: Ring d1600 у Geniled «под заказ», нужно подтвердить срок изготовления.</li>"
 "<li>Нужен сертификат пожарной безопасности Г1 В1 Д1 Т1 на панели, декларации ЕАЭС на всё, паспорта на русском.</li>"
 "</ul><h2>5. Что сделать до 18.11</h2><ul><li>Запросить КП: дилеры Ecophon (E-phon, ТК ОптСтрой) на 138 шт D1180 + D800/D600 и 663 подвеса; дилеры Geniled (Geniled Volga) на 35 Ring и комплект трека.</li>"
 "<li>Запросить ставку ТК Москва → Ессентуки (≈ 30 м³, тент 82 м³) и банк по гарантии.</li><li>Направить запрос разъяснений по D600, мощности Ring и цвету трековых.</li></ul>"
 "<div class=g style='margin-top:6px'>Источники цен (URL по каждой позиции) — лист «Источники» в Excel.</div></body></html>")
hp = f"{SCRATCH}/egp_econ.html"; open(hp, "w").write("".join(h))
subprocess.run(["/opt/pw-browsers/chromium", "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer", f"--print-to-pdf={os.path.abspath('../EGP_32616568269_EKONOMIKA.pdf')}", "file://" + hp], check=True, capture_output=True, timeout=180)
for sc, (rows, goods, cost) in RES.items(): print(sc, round(goods), round(cost), [round(NMC*(1-d)-cost) for d in DROPS])
