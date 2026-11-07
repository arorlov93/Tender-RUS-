"""Глубокий анализ цен закупки химии (07.11.2026): лестница цен завод → трейдер → дилер → импорт → контракты 2026, целевая цена закупки,
пересчёт маржи открытых лотов. Источники: data/companies/chem_deep_A.json, chem_deep_B.json (агенты, URL внутри), chem2_prices_A/B.json,
data/companies/chem_contracts_2026.json (контракты ЕИС, если выгружены). → ../HIMIYA_TSENY_ANALIZ_0711.xlsx + .pdf"""
import json, os, html, statistics, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from pdfutil import html_to_pdf
from tidy import tidy

D = "data/companies/"
# ── лестница цен, ₽/т с НДС (все цифры — из chem_deep_A/B и chem2_prices; «оценка» = расчёт, не котировка) ──
LAD = {
 "soda": dict(name="Сода кальцинированная марка Б", factory="не публикуют (БСК/Башхим — «по запросу»)", trader="РТХ 27 800", dealer="Аквахим 58 000 (мешки); saltrt 35 000",
              imp="Китай FOB 153–161 $/т → паритет с РФ", target=26000, tbasis="EXW Стерлитамак, вагон — оценка", conf="низкая", note="Контракты 2026: 34 650 (150 т), 44 460 (50 т), 46 000 (5 т) — победители продают по 35–46 тыс., значит их закупка ≈26–30 тыс. Цель подтверждается."),
 "natr46": dict(name="Натр едкий жидкий РД 46 % (за т раствора)", factory="не публикуют; ЕТК: FCA завода, цистерна 60 т, 100 % предоплата + возврат цистерн", trader="РТХ: Химпром 23 000, БСК 26 000",
                dealer="МК Магна 40 500; ОптХимТорг 42 000; Аквахим 53 000–55 000", imp="не применимо (жидкий)", target=21500, tbasis="FCA завода, цистерна — оценка", conf="средняя",
                note="Контракты 2026 смешаны по маркам: РД 53 700 (36 т), р-р 50 % 82 500, 20 % р-р 40 100. Дилеры в 1,8–2,4 раза дороже вагонного трейдера. В пересчёте на 100 % цель ≈46 700."),
 "natr100": dict(name="Натр едкий в пересчёте на 100 % NaOH", factory="—", trader="РТХ 55 800–56 500", dealer="—", imp="—", target=46700, tbasis="= 21 500 / 0,46 — оценка", conf="средняя", note="Проверять, в чём НМЦ: раствор или 100 %."),
 "natrS": dict(name="Натр едкий чешуированный", factory="не публикуют", trader="РТХ 58 500; Махаон ТД 61 000", dealer="МК Магна 73 000; Челхимпрод 72 000–78 000", imp="Китай FOB 410–440 $/т → 54 000–59 000",
               target=55000, tbasis="EXW Стерлитамак, от 20 т — оценка", conf="средняя", note="Контракт 2026: Ямал (303 т, северный завоз) 141 200 с доставкой. Импорт на уровне РТХ."),
 "gipo": dict(name="Гипохлорит натрия марки А", factory="не публикуют", trader="Технохимия «от 30 000»; МК Магна 41 000 (>10 т)", dealer="Аквахим 49 000 налив / 59 000–63 000 кубы", imp="не применимо",
              target=35000, tbasis="EXW налив; кубы/канистры +6 000–10 000", conf="низкая", note="Контракты 2026: 37 000 (57,6 т), 53 400–62 000 (5–58 т), 89 000 (3,4 т) — с доставкой. Короткий срок хранения — только ближний завод."),
 "chlor": dict(name="Хлор жидкий в контейнерах", factory="«по договорённости» (Каустик)", trader="нет", dealer="нет", imp="—", target=None, tbasis="—", conf="—",
               note="Посредников нет: ОПО, оборотные контейнеры. Не наш товар."),
 "sern": dict(name="Кислота серная техническая 1 сорт", factory="не публикуют (дефицит, ФАС против СУМЗ/Карабаш)", trader="ЯНОС 23 500 (22 500 от 20 т); Казцинк 34 160", dealer="МК Магна 34 000–37 700",
              imp="не применимо", target=28000, tbasis="EXW налив — оценка", conf="средняя", note="Контракты 2026 (вагоны 65 т, МО): июль 16 600–20 300 → август 28 700–36 400. Цена выросла почти вдвое за месяц (ограничение экспорта, дефицит серы). РТХ 13 000 — устаревшая."),
 "sernU": dict(name="Кислота серная улучшенная", factory="не найдено", trader="—", dealer="МК Магна 66 500", imp="—", target=55000, tbasis="≤55 000 — оценка", conf="низкая", note=""),
 "sa": dict(name="Сульфат алюминия твёрдый", factory="БИАКР ≈22 000 (НДС не подтверждён) ≈27 000 с НДС", trader="AquaTek 46 000–50 000", dealer="Химрус 66 000–81 000", imp="Китай ≈178 $/т",
            target=32000, tbasis="1 сорт МКР от 20 т (высший ≤45 000)", conf="низкая", note="Контракты 2026 (твёрдый, 20–60 т): 45 000–62 000, Татарстан 85 000; 300 т — 18 500 (вероятно жидкий). Дилеры в 2,5–3 раза дороже производителя."),
 "saL": dict(name="Сульфат алюминия жидкий", factory="не найдено", trader="—", dealer="Химрус 40 000–55 000 (мелкая тара)", imp="—", target=20000, tbasis="≤20 000 EXW — расчёт по Al2O3", conf="низкая", note=""),
 "tio": dict(name="Натрия тиосульфат", factory="производителей в РФ не найдено", trader="—", dealer="РТХ 99 000; AquaTek 114 000", imp="Китай 350 $/т FOB → ≈47 000 landed", target=55000,
             tbasis="склад импортёра — по контракту 2026", conf="средняя", note="Контракт 2026 (Рязань, 22 т): 56 000 ₽/т — поставщик продаёт ниже, чем РТХ покупателю (99 000). Дилер ≈2× к импорту."),
 "tnf": dict(name="Тринатрийфосфат", factory="Ветлужский ХЗ через ТД 116 000–118 000", trader="—", dealer="chem-ex 145 000; Химреактив 195 000", imp="Китай 3 140 CNY/т", target=90000,
             tbasis="ниже цен победителей 2026", conf="средняя", note="Контракты 2026 (12 шт.): 96 700–129 000, Форвард — 96 720 с доставкой в Тюмень/Челябинск. Победители продают ниже прайса ТД Ветлуги (116 000) → реальная закупка ≤90 000."),
 "izv": dict(name="Известь негашёная комовая", factory="АТОМ (Богданович) 9 000 навал / 10 500 МКР", trader="ural-izvest 9 000–10 000", dealer="AquaTek 16 000–22 000", imp="—", target=9000,
             tbasis="EXW, завод ≤300 км от заказчика", conf="средняя", note="Решает доставка."),
 "limon": dict(name="Кислота лимонная", factory="«Органические кислоты», «Цитрон» — цен нет", trader="—", dealer="РТХ 96 000; МК Магна 99 000", imp="Китай 6 200 CNY/т → ≈72 000 landed", target=82000,
               tbasis="склад импортёра", conf="средняя", note="ЕЭК: антидемпинг против Китая до 58,8 %."),
 "ntf": dict(name="Кислота НТФ (ATMP 95 %), ₽/кг", factory="Химпром — производство не подтверждено", trader="—", dealer="Татсорб 354–364; Пульс цен 350; Химрус 290–300", imp="Китай 10 500 CNY/т → ≈160 ₽/кг landed",
             target=210, tbasis="склад импортёра, ₽/кг — оценка", conf="низкая", note="Дилеры в 1,8–2,3 раза дороже импорта."),
}
# ── лоты: (ключ товара, НМЦ за ед. с НДС, кол-во, доставка за ед., доп. тара за ед., НМЦ с/без НДС, ед.) ──
LOTS = [
 ("32616546242", "Новгородский водоканал — сода кальц. Б", "23.11", "soda", 56407, 834.1, 6000, 0, "с", "т"),
 ("32616531354", "ВТЕ Юго-Восток (Москва) — сода кальц.", "17.11", "soda", 45660, 3078, 5000, 0, "с", "т"),
 ("32616550391", "БГК Уфа — сода кальц.", "16.11", "soda", 45060, 50.4, 1500, 0, "с", "т"),
 ("32616474821", "СКАТЗ — натр едкий РД (НМЦ ≈ 100 %)", "13.11", "natr100", 81679, 75, 3000, 0, "с", "т"),
 ("32616541193", "Зуевская ТЭС — натр едкий, если НМЦ за раствор 46 %", "19.11", "natr46", 44346, 2511, 5000, 0, "с", "т"),
 ("32616541193b", "Зуевская ТЭС — натр едкий, если НМЦ за 100 %", "19.11", "natr100", 44346, 2511, 5000, 0, "с", "т"),
 ("32616534976", "Т Плюс Саратов — тиосульфат натрия", "12.11", "tio", 145863, 10, 3000, 0, "с", "т"),
 ("32616541289", "РМ Нанотех — НТФ 95 % (₽/кг)", "13.11", "ntf", 321.13, 14000, 5, 0, "с", "кг"),
 ("32616543663", "Тверь Водоканал — сульфат алюминия (НМЦ без НДС)", "16.11", "sa", 139945, 18, 5400, 13000, "без", "т"),
 ("32616523021", "ЯВК Ярославль — сульфат алюминия жидкий", "16.11", "saL", 27400, 6000, 1000, 0, "с", "т"),
 ("32616546151", "Водоканал-Сервис (МО) — гипохлорит (тара не указана)", "23.11", "gipo", 57000, 300, 4800, 0, "с", "т"),
 ("32616549952", "Леноблводоканал — гипохлорит в канистрах", "23.11", "gipo", 92250, 607.2, 9300, 10000, "с", "т"),
 ("32616547400", "УЛВК — гипохлорит", "23.11", "gipo", 82533, 70, 9300, 8000, "неясно", "т"),
 ("32616552377", "КВГ Геленджик — гипохлорит", "16.11", "gipo", 71000, 56.3, 11400, 0, "неясно", "т"),
 ("32616549809", "Леноблводоканал — гипохлорит в кубах", "16.11", "gipo", 51832, 190.5, 9300, 8000, "с", "т"),
 ("32616543474", "Усть-Вымская ТК — гипохлорит", "12.11", "gipo", 45140, 120, 7200, 0, "с", "т"),
 ("32616526278", "Водоканал Рязани — гипохлорит", "20.11", "gipo", 42360, 1449, 4500, 0, "с", "т"),
 ("32616524946", "ВТЕ Юго-Восток — серная кислота 27 тыс. т", "16.11", "sern", 33762, 27120, 3000, 0, "с", "т"),
 ("32616540560", "Омск — серная улучшенная", "12.11", "sernU", 70950, 24, 12000, 0, "с", "т"),
]
DEALER = {}  # прежняя (дилерская) закупка из chem2_prices
for f in ("chem2_prices_A.json", "chem2_prices_B.json"):
    for x in json.load(open(D + f)): DEALER[str(x["n"])] = x.get("buy_unit_vat")

def margin(nu, vat, cost_v, q):
    if vat == "без": bid = nu * .9; usn = (bid - cost_v) / bid; osno = (bid - cost_v / 1.22) / bid; rev = bid
    else:
        bid_v = nu * .9; bid_nv = bid_v / 1.22
        usn = ((bid_nv - cost_v) / bid_nv) if vat == "неясно" else (bid_v - cost_v) / bid_v; osno = (bid_nv - cost_v / 1.22) / bid_nv; rev = bid_nv if vat == "неясно" else bid_v
    return usn, osno, usn * rev * q

# ── контракты ЕИС 2026 ──
import re
RAW = json.load(open(D + "chem_contracts_2026_raw.json")) if os.path.exists(D + "chem_contracts_2026_raw.json") else {}
BADN = re.compile(r"мыло|бумаг|губк|паст[аы]|насос|водосчетч|контейнер|метотрексат|услуг|порошок чистящ", re.I)
CT = {}
for k, L in RAW.items():
    for r in L:
        if BADN.search(r["name"]) or not r.get("unit_t") or r["unit_t"] > 400000 or (r.get("qty") or 0) < 1: continue  # розница <1 т и ошибки ед. изм. — вне выборки
        CT.setdefault(k, []).append(r)
json.dump(CT, open(D + "chem_contracts_2026.json", "w"), ensure_ascii=False, indent=1)
CK = {"soda": "soda", "natr46": "natr", "natr100": "natr", "natrS": "natr", "gipo": "gipo", "chlor": "chlor", "sern": "sern", "sernU": "sern", "sa": "sa", "saL": "sa", "tio": "tio", "tnf": "tnf", "izv": "izv", "limon": "limon", "ntf": None}
CF = {"natr46": r"р-р|раствор|РД", "natrS": r"чешуир|твёрд|ТР", "natr100": r"^$", "sernU": r"^$", "saL": r"^$"}
def cstat(k):
    xs = [r["unit_t"] for r in CT.get(CK.get(k) or "", []) if r.get("unit_t") and (k not in CF or re.search(CF[k], r["name"] + " " + r.get("subject", ""), re.I))]
    if not xs: return None
    return dict(n=len(xs), med=statistics.median(xs), lo=min(xs), hi=max(xs))

rows = []
for n, name, close, k, nu, q, dl, tara, vat, unit in LOTS:
    L = LAD[k]; t = L["target"]
    if t is None: continue
    ut, ot, pt = margin(nu, vat, t + dl + tara, q)
    dv = DEALER.get(n[:11])
    ud = margin(nu, vat, dv + dl + tara, q)[0] if dv else None
    dec = "ИДЁМ" if ut >= .15 else "ТОНКО" if ut >= .05 else "НЕ ИДЁМ"
    if n.endswith("b") or n == "32616541193": dec = "ПРОВЕРИТЬ ТЗ: " + dec
    if n == "32616543663": dec = "ПРОВЕРИТЬ ТЗ: " + dec
    rows.append(dict(n=n[:11], name=name, close=close, k=k, nu=nu, q=q, unit=unit, dl=dl, tara=tara, vat=vat, t=t, dv=dv, ud=ud, ut=ut, ot=ot, pt=pt, dec=dec, conf=L["conf"]))
rows.sort(key=lambda r: (0 if r["dec"].startswith("ИДЁМ") else 1 if "ИДЁМ" in r["dec"] or r["dec"].startswith("ТОНКО") else 2, -r["pt"]))

# ── Excel ──
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Лестница цен"
ws.append(["Товар", "Завод / ТД завода", "Крупный трейдер", "Дилеры", "Импорт", "Контракты ЕИС 2026: медиана ₽/т (мин–макс, n)", "ЦЕЛЕВАЯ закупка ₽/т с НДС", "База", "Уверенность", "Вывод"])
for k, L in LAD.items():
    c = cstat(k)
    ws.append([L["name"], L["factory"], L["trader"], L["dealer"], L["imp"], (f"{c['med']:,.0f} ({c['lo']:,.0f}–{c['hi']:,.0f}, n={c['n']})".replace(",", " ") if c else "нет данных"), L["target"], L["tbasis"], L["conf"], L["note"]])
s2 = wb.create_sheet("Маржа лотов по целевой цене")
s2.append(["Решение", "№ ЕИС", "Лот", "Подача", "НМЦ за ед.", "НМЦ с/без НДС", "Кол-во", "Ед.", "Закупка ДИЛЕР", "Закупка ЦЕЛЬ", "Доставка", "Тара", "Контракты ЕИС 2026 по товару (медиана, n)", "УСН −10 % (дилер)", "УСН −10 % (цель)", "ОСНО −10 % (цель)", "Прибыль УСН −10 % (цель), ₽", "Уверенность"])
for r in rows:
    s2.append([r["dec"], r["n"], r["name"], r["close"], r["nu"], r["vat"], r["q"], r["unit"], r["dv"], r["t"], r["dl"], r["tara"], (f"{cstat(r['k'])['med']:,.0f} (n={cstat(r['k'])['n']})".replace(",", " ") if cstat(r['k']) else "—"), r["ud"], r["ut"], r["ot"], r["pt"], r["conf"]])
    X = s2.max_row
    for j in (14, 15, 16): s2.cell(X, j).number_format = "0%"
    for j in (5, 9, 10, 11, 12, 17): s2.cell(X, j).number_format = "#,##0"
    s2.cell(X, 1).fill = PatternFill("solid", fgColor="c6efce" if r["ut"] >= .15 else "fff2cc" if r["ut"] >= .05 else "f8cbad")
if CT:
    s3 = wb.create_sheet("Контракты ЕИС 2026"); s3.append(["Товар", "Реестровый №", "Закон", "Дата", "Регион", "Позиция", "Цена за ед.", "Ед.", "₽/т", "Кол-во", "Сумма контракта", "Поставщик ИНН"])
    for k, L in CT.items():
        for r in L: s3.append([k, r.get("reg"), r.get("law"), r.get("date"), r.get("region"), r.get("name"), r.get("price"), r.get("okei"), r.get("unit_t"), r.get("qty"), r.get("sum"), r.get("supplier")])
for s in wb:
    for c in s[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
for i, w in enumerate([30, 32, 30, 32, 28, 26, 14, 26, 10, 40], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
for i, w in enumerate([18, 13, 42, 8, 11, 9, 9, 5, 11, 11, 9, 9, 10, 10, 10, 14, 10], 1): s2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
tidy(wb, row_h=45); wb.save("../HIMIYA_TSENY_ANALIZ_0711.xlsx")

# ── PDF (2 стр.) ──
e = lambda s: html.escape(str(s if s not in (None, "") else "—"))
fm = lambda v: f"{v:,.0f}".replace(",", " ") if isinstance(v, (int, float)) else "—"
pc = lambda v: f"{v*100:.0f} %" if isinstance(v, (int, float)) else "—"
cl = lambda v: "n" if not isinstance(v, (int, float)) else "ok" if v >= .15 else "warn" if v >= .05 else "bad"
css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:7.3px;color:#111;margin:0}h1{font-size:12px;margin:0 0 3px}h2{font-size:10px;margin:5px 0 2px;border-bottom:2px solid #305496}table{border-collapse:collapse;width:100%}
th,td{border:1px solid #bbb;padding:1px 3px;vertical-align:top}th{background:#305496;color:#fff;text-align:left}td.n{text-align:right;white-space:nowrap}.ok{background:#c6efce}.warn{background:#fff2cc}.bad{background:#f8cbad}.box{border:1px solid #305496;padding:4px 6px;background:#f3f6fb;margin:3px 0}.s{font-size:6.6px;color:#444}.pb{page-break-before:always}"""
go = [r for r in rows if r["dec"] == "ИДЁМ"]
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body><h1>Химия: реальные цены закупки — лестница «завод → трейдер → дилер → импорт → контракты» (07.11.2026)</h1>",
     "<div class='box'><b>Главное.</b> Заводы (БСК/Башхим, Каустик, Химпром, ФосАгро, ЕвроХим, УГМК, Аурат) цены не публикуют — только по запросу, вагонными нормами и чаще по предоплате. "
     "Нижняя публичная планка — вагонные трейдеры (РТХ, ЯНОС через Химспециализацию, ТД Ветлужского завода): по оценке это завод + 5–10 %. Мелкие дилеры (Аквахим, МК Магна, AquaTek) дороже в 1,5–3 раза — по ним считать нельзя. "
     "Целевая цена = ниже трейдера на 3–8 % (вагон/цистерна), для импортных товаров — расчёт «Китай FOB + логистика + маржа импортёра». "
     "Проверка фактом: 61 контракт ЕИС 2026 (44/223-ФЗ, партии ≥1 т) — цены, по которым победители реально продают; закупка у них ниже. Без КП это ОЦЕНКИ; подтвердить — запросами на вагонные партии в ТД БСК (sodastr@soda.ru), ЕТК, Химпром, Каустик, ЯНОС, Ветлужский ХЗ.</div>",
     "<h2>1. Лестница цен, ₽/т с НДС</h2><table><tr><th>Товар</th><th>Завод / ТД</th><th>Трейдер</th><th>Дилеры</th><th>Импорт</th><th>Контракты ЕИС 2026</th><th>ЦЕЛЬ</th><th>Ув.</th><th>Вывод</th></tr>"]
for k, L in LAD.items():
    c = cstat(k)
    h.append(f"<tr><td><b>{e(L['name'])}</b></td><td class=s>{e(L['factory'])}</td><td class=s>{e(L['trader'])}</td><td class=s>{e(L['dealer'])}</td><td class=s>{e(L['imp'])}</td>"
             f"<td class=s>{(fm(c['med']) + ' (' + fm(c['lo']) + '–' + fm(c['hi']) + ', n=' + str(c['n']) + ')') if c else 'нет данных'}</td><td class=n><b>{fm(L['target'])}</b><br><span class=s>{e(L['tbasis'])}</span></td><td class=s>{e(L['conf'])}</td><td class=s>{e(L['note'])}</td></tr>")
h.append("</table>")
h.append(f"<h2 class=pb>2. Маржа открытых лотов: по дилерской цене (как раньше) и по целевой</h2><div class='box'>Подаём ({len(go)}): маржа ≥15 % на УСН при −10 % по целевой цене; прибыль ≈{sum(r['pt'] for r in go)/1e6:.1f} млн ₽. "
         "Тара: канистры/кубы +8 000–13 000 ₽/т к наливу. Доставка — оценка по расстоянию.</div><table><tr><th>Решение</th><th>№ / подача</th><th>Лот</th><th>НМЦ за ед.</th><th>Кол-во</th><th>Дилер → цель</th><th>Контракты 2026, медиана</th><th>Доставка + тара</th><th>УСН (дилер)</th><th>УСН (цель)</th><th>ОСНО (цель)</th><th>Прибыль УСН, млн</th></tr>")
for r in rows:
    h.append(f"<tr><td class='{cl(r['ut'])}'><b>{e(r['dec'])}</b></td><td class=s>{r['n']}<br>{e(r['close'])}</td><td>{e(r['name'])}</td><td class=n>{fm(r['nu'])}<br><span class=s>{e(r['vat'])} НДС</span></td><td class=n>{fm(r['q'])} {r['unit']}</td>"
             f"<td class=n>{fm(r['dv'])} → <b>{fm(r['t'])}</b></td><td class=n>{(fm(cstat(r['k'])['med']) + ' (n=' + str(cstat(r['k'])['n']) + ')') if cstat(r['k']) else '—'}</td><td class=n>{fm(r['dl'])} + {fm(r['tara'])}</td><td class='n {cl(r['ud'])}'>{pc(r['ud'])}</td><td class='n {cl(r['ut'])}'><b>{pc(r['ut'])}</b></td><td class='n {cl(r['ot'])}'>{pc(r['ot'])}</td><td class=n>{r['pt']/1e6:.2f}</td></tr>")
h.append("</table></body></html>")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/chem_deep.html"
open(hp, "w").write("".join(h)); html_to_pdf(hp, "../HIMIYA_TSENY_ANALIZ_0711.pdf")
for r in rows: print(r["dec"], r["n"], r["name"][:40], pc(r["ud"]), "->", pc(r["ut"]), round(r["pt"] / 1e6, 2))
