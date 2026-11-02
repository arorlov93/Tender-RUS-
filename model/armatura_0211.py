"""Партия закупок 02.11: запорная арматура (задвижки, клапаны, краны) — сравнение НМЦ с ценами полученных КП.
База КП — все КП по Камчатке (ДН.РУ, Теплогранд/LD, БАЗ № БАЗ2026-00077, АДЛ, КИПИКА, Лидер-Энерго, Тула, ЭПАС) с НМЦ ед. Камчатки,
сопоставление по ключу «марка + Ду + Ру». ДГК 32616403815 — ТЗ (1467 поз.) в data/dgk; Сахалин 11201 — НМЦ ед. в data/sakh/sakh_prices.json.
Цены поставщиков — с НДС (мы на УСН). → ../ARMATURA_KP_0211.xlsx"""
import json, re, runpy, io, contextlib, collections, openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from tidy import tidy

with contextlib.redirect_stdout(io.StringIO()):
    VY = runpy.run_path("kamch_vygodno.py")
FM = json.load(open("data/kamch/forms/kp_form_map.json"))
MK = re.compile(r"\b(\d{2}[а-яa-z]{1,3}\d{1,3}[а-яa-z]{0,3}\d{0,2})\b", re.I)


def key(name):
    n = name.replace("Ду", "DN ").replace("Ру", "PN ")
    m = MK.search(n); dn = re.search(r"DN\s*0*(\d+)", n); pn = re.search(r"PN\s*(\d+(?:[.,]\d+)?)", n)
    if not (m and dn and pn): return None
    p = float(pn.group(1).replace(",", ".")); p = p * 10 if p < 1 else p
    return (m.group(1).lower(), int(dn.group(1)), round(p))


SRC = ("ДН.РУ", "Теплогранд (LD)", "БАЗ", "ТД АДЛ", "КИПИКА", "Лидер-Энерго")
KP = {}  # key -> {src: (цена ед. с НДС, соответствие ТЗ)}, + НМЦ ед. Камчатки
for r in VY["res"]:
    k = key(r["name"])
    if not k: continue
    e = KP.setdefault(k, {"tt": r["t"], "kam_nmc": FM[r["t"]]["nmc"], "kam_name": r["name"]})
    for o in r["offs"]:
        s = next((x for x in SRC if o["sup"].startswith(x)), None)
        if s and (s not in e or o["c"] < e[s][0]): e[s] = (round(o["c"], 2), o["v"])

thin = Side(style="thin", color="BFBFBF"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin); M = "#,##0"
def hd(ws):
    for c in ws[ws.max_row]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True, vertical="center")
wb = openpyxl.Workbook()

# ---------- ДГК 32616403815
U = json.load(open("data/dgk/dgk_unique.json"))
WEB = {}
for f in ("price_zadv", "price_kran", "price_klap", "price_other"):
    for x in json.load(open(f"data/dgk/prices/{f}.json")):
        if x.get("price"): WEB[x["key"]] = x["price"]
wd = wb.active; wd.title = "ДГК — наши КП"
wd.append(["АО «ДГК» 32616403815 — арматура среднего и низкого давления (НМЦК 186 531 454,93 ₽; 1467 позиций / 413 уникальных). Позиции, по которым у нас есть КП (совпадение марка+Ду+Ру с Камчаткой)"]); wd["A1"].font = Font(bold=True, size=12)
wd.append(["Позиция ДГК", "Кол-во", "ДН.РУ, ₽/шт", "Теплогранд (LD), ₽/шт", "БАЗ, ₽/шт", "Мин. по КП, ₽/шт", "Кто дешевле", "Соответствие ТЗ (по Камчатке)", "НМЦ ед. Камчатки (тот же товар), ₽", "Мин. КП / НМЦ Камчатки", "Сумма по мин. КП, ₽", "Сумма по НМЦ Камчатки, ₽", "Аналог в ТЗ Камчатки"]); hd(wd)
T = collections.Counter(); n_m = 0
for u in U:
    k = key(u["name"] + " " + (u.get("spec") or ""))
    e = KP.get(k) if k else None
    if not e: continue
    pr = {s: e[s] for s in SRC if s in e}
    if not pr: continue
    n_m += 1; q = u["qty"]
    best = min(pr.items(), key=lambda x: x[1][0])
    wd.append([u["name"], q, (e.get("ДН.РУ") or (None,))[0], (e.get("Теплогранд (LD)") or (None,))[0], (e.get("БАЗ") or (None,))[0], best[1][0], best[0], best[1][1], e["kam_nmc"],
               best[1][0] / e["kam_nmc"] if e["kam_nmc"] else None, best[1][0] * q, e["kam_nmc"] * q, f'{e["tt"]}: {e["kam_name"]}'])
    X = wd.max_row
    for j in (3, 4, 5, 6, 9, 11, 12): wd.cell(X, j).number_format = M
    wd.cell(X, 10).number_format = "0%"; r_ = wd.cell(X, 10).value or 0
    wd.cell(X, 10).fill = PatternFill("solid", fgColor="C6EFCE" if r_ < 0.75 else "FFF2CC" if r_ < 1 else "F8CBAD")
    T["kp"] += best[1][0] * q; T["nmc"] += e["kam_nmc"] * q
    for s in ("ДН.РУ", "Теплогранд (LD)", "БАЗ"):
        if s in pr: T[s + "_n"] += 1
wd.append(["ИТОГО по совпавшим", n_m, "", "", "", "", "", "", "", T["kp"] / T["nmc"] if T["nmc"] else None, T["kp"], T["nmc"]])
X = wd.max_row; [setattr(c, "font", Font(bold=True)) for c in wd[X]]; wd.cell(X, 10).number_format = "0%"; wd.cell(X, 11).number_format = M; wd.cell(X, 12).number_format = M
wd.append([]); wd.append(["Ранее (22.10) по всем 1467 позициям с ценами рынка 2026: себестоимость ~204 млн ₽ > НМЦК 186,5 млн ₽ — лот убыточный (DGK_32616403815.xlsx). Задвижки дают 136 млн себестоимости из 186 млн НМЦК."])
for i, w in enumerate([46, 8, 13, 13, 13, 13, 16, 18, 14, 11, 15, 15, 45], 1): wd.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
wd.freeze_panes = "C3"
dgk = dict(n=n_m, kp=T["kp"], nmc=T["nmc"])

# ---------- Сахалин 11201 (лот 1): НМЦ ед. есть в ТТ
S = json.load(open("data/sakh/sakh_prices.json"))["11201"]
ws = wb.create_sheet("Сахалин лот 1 — НМЦ vs КП")
ws.append(["ПАО «Сахалинэнерго» лот 1 № 11201 (НМЦ 28,42 млн ₽, подача 05.11): НМЦ ед. vs себестоимость (КП Сахалина/рынок) и наши КП по Камчатке"]); ws["A1"].font = Font(bold=True, size=12)
ws.append(["Партия/№", "Позиция", "Кол-во", "НМЦ ед., ₽", "Себес ед. (прежний расчёт), ₽", "Источник", "Мин. КП по Камчатке, ₽/шт", "Кто", "Лучшая себес ед., ₽", "Маржа ед.", "НМЦ сумма, ₽", "Себес сумма, ₽"]); hd(ws)
TS = collections.Counter()
for x in S:
    k = key(x["name"]); e = KP.get(k) if k else None
    kp = min(((v[0], s) for s, v in e.items() if s in SRC), default=None) if e else None
    c0 = x.get("cost_unit"); c = min([v for v in (c0, kp[0] if kp else None) if v] or [None]) if (c0 or kp) else None
    ws.append([f'{x["part"]}/{x["n"]}', x["name"], x["qty"], x["nmc_unit"], c0, str(x.get("src") or "")[:60], kp[0] if kp else None, kp[1] if kp else "", c,
               (x["nmc_unit"] - c) / x["nmc_unit"] if c and x["nmc_unit"] else None, x["nmc_unit"] * x["qty"], (c or 0) * x["qty"]])
    X = ws.max_row
    for j in (4, 5, 7, 9, 11, 12): ws.cell(X, j).number_format = M
    ws.cell(X, 10).number_format = "0%"
    mg = ws.cell(X, 10).value
    if mg is not None: ws.cell(X, 10).fill = PatternFill("solid", fgColor="C6EFCE" if mg >= 0.2 else "FFF2CC" if mg >= 0 else "F8CBAD")
    TS["nmc"] += x["nmc_unit"] * x["qty"]; TS["cost"] += (c or 0) * x["qty"]; TS["nocost"] += 0 if c else 1; TS["kp_better"] += 1 if kp and (not c0 or kp[0] < c0) else 0
ws.append(["ИТОГО", "", "", "", "", "", "", "", "", (TS["nmc"] - TS["cost"]) / TS["nmc"], TS["nmc"], TS["cost"]])
X = ws.max_row; [setattr(c, "font", Font(bold=True)) for c in ws[X]]; ws.cell(X, 10).number_format = "0%"; ws.cell(X, 11).number_format = M; ws.cell(X, 12).number_format = M
ws.append([f"Позиций без цены: {TS['nocost']}; КП по Камчатке дешевле прежней себестоимости по {TS['kp_better']} позициям. Доставка до Сахалина в себестоимость не включена."])
for i, w in enumerate([9, 46, 7, 12, 14, 30, 14, 14, 13, 9, 14, 14], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "C3"
sakh = dict(nmc=TS["nmc"], cost=TS["cost"], nocost=TS["nocost"], better=TS["kp_better"])

# ---------- сводка по всем закупкам арматуры партии
sv = wb.create_sheet("Сводка", 0)
sv.append(["Партия 02.11: запорная арматура (задвижки, клапаны, краны) — НМЦ против наших КП"]); sv["A1"].font = Font(bold=True, size=12)
sv.append(["№ закупки", "Заказчик / предмет", "НМЦ, ₽", "Подача", "Что есть", "Результат"]); hd(sv)
LIST = [
 ("32616403815", "АО «ДГК» — арматура среднего и низкого давления", 186531454.93, "05.11 17:00", f"ТЗ 1467 поз.; совпало с нашими КП {dgk['n']} уникальных поз.",
  f"По совпавшим: мин. КП = {dgk['kp'] / dgk['nmc']:.0%} от НМЦ Камчатки на тот же товар. По всему лоту себестоимость ~204 млн > НМЦК 186,5 млн — НЕ ИДЁМ (задвижки российских заводов, БАЗ в 2–3 раза дороже)."),
 ("32616494868", "ПАО «Сахалинэнерго» — лот 1 (11201)", 28420000, "05.11", f"НМЦ ед. по 83 поз.; КП 19 поз. + КП Камчатки",
  f"Себестоимость по лучшим ценам {sakh['cost'] / 1e6:.2f} млн при НМЦ {sakh['nmc'] / 1e6:.2f} млн (маржа {(sakh['nmc'] - sakh['cost']) / sakh['nmc']:.0%} до доставки, без цены {sakh['nocost']} поз.) — см. лист."),
 ("32616358208", "АО «ДГК» — арматура высокого давления", 373625078.67, "11.11", "ТЗ нет (энергоарматура ЧЗЭМ/ЗиО)", "Наших КП на высокое давление нет; требование паспортов на ранее поставленную продукцию — НЕ ИДЁМ."),
 ("32616420912", "ПАО «ТГК-1» — общепромышленная арматура", 1226100000, "09.11", "ТЗ нет", "Нужен ТЗ; по номенклатуре 30с41нж/30ч6бр/шаровые сравним с этой базой."),
 ("8235", "РМЗ (Башкортостан) — запорная арматура", 7432559, "11.11", "ТЗ нет", "Нужен ТЗ; БАЗ рядом, но его цены высокие (КП БАЗ: 30с41нж Ду50 — 30 750 ₽ без НДС); ориентир — LD/ДН.РУ."),
 ("32616465624", "УПТЖ для ППД (Татнефть)", 1719646.12, "10.11", "ТЗ (Excel) на ЭТП", "Нужен ТЗ."),
 ("1038631", "ТСК Мосэнерго — чугунная арматура", 463478.73, "13.11", "ТЗ нет", "Нужен ТЗ; 30ч6бр есть у ЭКС-ФОРМА/ПИК АЛГРУПП (реестр)."),
 ("1029143", "Т Плюс Пенза — затворы дисковые", 300000, "03.11 (прошла)", "—", "Срок прошёл."),
]
for r in LIST:
    sv.append(list(r)); X = sv.max_row; sv.cell(X, 3).number_format = M
    for c in sv[X]: c.border = BOX; c.alignment = Alignment(wrap_text=True, vertical="top")
sv.append([]); sv.append(["Порядок цен по нашим КП (с НДС, за шт): задвижка 30с41нж Ду50 Ру16 — ДН.РУ ~8 тыс., БАЗ 37,5 тыс.; Ду100 — 14,5 тыс. / 60,6 тыс.; Ду150 — 25,8 тыс. / 112 тыс.; 30с64нж Ду100 Ру25 — ДН.РУ 75,9 тыс., БАЗ 70,6 тыс. Краны LD (Теплогранд) — 0,25–49 тыс."])
for i, w in enumerate([18, 44, 16, 14, 36, 80], 1): sv.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
tidy(wb, row_h=30)
wb.save("../ARMATURA_KP_0211.xlsx")
print("ДГК совпало", dgk, "Сахалин", sakh)
