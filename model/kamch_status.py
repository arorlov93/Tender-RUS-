"""Камчатка 32616346943 — сводный статус закупки по всем 315 позициям: обновлённые счета ДН.РУ 819892/819896 (26.10), КП Тулаэлектропривода, КП Лидер-Энерго,
проверенные заводы по оставшимся позициям. Решение по каждой позиции: брать / брать при условии / торг (наш бюджет) / у других / не отработано."""
import json, re, openpyxl
from collections import defaultdict
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from tidy import tidy

TORG = 0.10
thin = Side(style="thin", color="BFBFBF"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin); M = '#,##0" ₽"'; P = "+0%;-0%;0%"
TT = {(t["table"], t["n_table"]): t for t in json.load(open("data/kamch/tz_verbatim.json"))}
CN = re.compile(r"Китай|BasicStandard|Benarmo|Dendor|Мегастрой", re.I)
MARK = re.compile(r"\b(\d{2}[а-яa-z]{1,2}\d{1,3}[а-яa-z]{0,3}\d{0,2}|VFG-2R|VFM|AFD|NVD|РД-НО|УРРД|ПСК|ВДХ|АРТ)", re.I)
dn = lambda s: (lambda m: int(m.group(1)) if m else None)(re.search(r"(?:Ду|DN)\s?(\d+)", s))
pn = lambda s: (lambda m: float(m.group(1).replace(",", ".")) if m else None)(re.search(r"(?:Ру|PN)\s?(\d+(?:[.,]\d+)?)", s))


def load_inv(f):
    d = json.load(open(f)); g = defaultdict(list)
    for x in d["items"]: g[x["z"]].append(x)
    return g


def old_unit(f):  # первая версия счёта — для колонки «было»
    try:
        d = json.load(open(f)); g = defaultdict(float)
        for x in d["items"]: g[x["z"]] += x.get("sum_vat") or x.get("price_vat") or 0
        return g
    except FileNotFoundError: return {}


INV = {"Задвижки": load_inv("data/kamch/kp/dnru_819892_v2.json"), "Клапаны": load_inv("data/kamch/kp/dnru_819896_v2.json")}
OLD = {"Задвижки": old_unit("data/kamch/kp/dnru_819892.json"), "Клапаны": old_unit("data/kamch/kp/dnru_819896.json")}
TULA = json.load(open("data/kamch/kp/tula_10-160.json")); LIDER = {x["tt"]: x for x in json.load(open("data/kamch/kp/lider-energo_2022.json"))["items"]}
ADL = {x["tt"]: x for x in json.load(open("data/kamch/kp/adl_utd112759.json"))["items"]}
KIP = {x["tt"]: x for x in json.load(open("data/kamch/kp/kipika_0110.json"))["items"]}
# ДН.РУ: позиции «в работе» (письмо 26.10) и вопросы по Kvs
PEND = {"Задвижки": {2: "ДН.РУ подбирает"}, "Клапаны": {4: "ДН.РУ подбирает", 9: "ДН.РУ подбирает", 10: "ДН.РУ подбирает", 11: "ДН.РУ подбирает", 20: "ДН.РУ подбирает",
        21: "ДН.РУ подбирает", 58: "ДН.РУ подбирает (Dungs — импорт, нужен рос. аналог)", 68: "ДН.РУ подбирает", 69: "ДН.РУ подбирает",
        94: "ДН.РУ ждёт Kvs — ответ из ТЗ: 4 м³/ч", 95: "ДН.РУ ждёт Kvs — ответ из ТЗ: 4 м³/ч", 96: "ДН.РУ ждёт Kvs — ответ из ТЗ: макс. 65 м³/ч, допустимо 40 м³/ч",
        97: "ДН.РУ ждёт Kvs — ответ из ТЗ: 4 м³/ч (для Ду80 — подтвердить у заказчика)"}}
VER = [x for f in ("data/kamch/rem_krany.json", "data/kamch/rem_zatvory.json") for x in json.load(open(f)) if x.get("verified")]
fact = lambda tt: "; ".join(f'{x["company"].split("(")[0].strip()} {str(x.get("phone") or "").split(";")[0]}' for x in VER if tt in (x.get("covers_tt") or []))
SUP = json.load(open("data/kamch/suppliers.json"))
SKIP = re.compile(r"не использовать|не опубликован|только через дилеров|консультация|ДН\.ру|DN\.ru|ТехГидроМаш", re.I)
KEY = [(r"30нж|нерж.*задв", "Задвижки нерж"), (r"30ч|чугун.*задв", "Задвижки чугун"), (r"задвижк|30с", "Задвижки сталь"), (r"обратн|NVD|Dendor|16ч|19[чсн]|16нж", "Клапаны обратн"),
       (r"предохран|17[лс]|ПСК|ARI-SAFE|СППК", "Клапаны предохран"), (r"регулир|регулятор|VFG|AFD|РД-НО|АРТ|25ч|25с|Dungs|газов", "Клапаны регулир"), (r"клапан|вентил|ВДХ|15|14нж|22б", "Клапаны запорн")]


def alt(name):
    k = next((v for rx, v in KEY if re.search(rx, name, re.I)), None)
    g = next((g for g in SUP if k and g["group"].lower().startswith(k.lower())), None)
    return "; ".join(f'{s["company"].split("(")[0].strip()} {str(s.get("phone") or "").split(";")[0].split(chr(10))[0]}' for s in (g or {}).get("suppliers", []) if not SKIP.search(str(s.get("checked")) + s["company"]))


wb0 = openpyxl.load_workbook("../KAMCHATKA_BY_GROUP_KP.xlsx"); rows = []
for g in ("Задвижки", "Клапаны", "Краны", "Затворы", "Приводы и прочее"):
    for r in wb0[g].iter_rows(values_only=True):
        if not (r[0] and str(r[0]).isdigit() and re.match(r"^\d\.\d+ / \d+", str(r[1] or ""))): continue
        t, n = [x.strip() for x in r[1].split("/")]; z = int(r[0]); tz = TT[(t, int(n))]
        d = dict(grp=g, z=z, tt=f"{t}/{n}", name=tz["Наименование"], q=float(tz["qty"]), our=float(r[8]), inv=None, old=None, sup="", offer="", srok="", flags=[])
        lines = INV.get(g, {}).get(z)
        if lines:
            d.update(inv=sum(x["sum_vat"] for x in lines), sup="ДН.РУ", srok=lines[0]["srok"], old=OLD[g].get(z),
                     offer=lines[0]["name"] + (f" + КОФ ({len(lines) - 1} стр.)" if len(lines) > 1 else ""))
            main, nm = lines[0]["name"], tz["Наименование"]
            mk = [m.lower() for m in MARK.findall(nm)]
            if "Задвижка" in main and "задвижк" not in nm.lower(): d["flags"].append("BAD:другой товар (задвижка вместо клапана)")
            elif mk and not any(m in main.lower() for m in mk): d["flags"].append(f"COND:аналог вместо {mk[0]} — обоснование эквивалентности")
            if dn(nm) and dn(main) and dn(nm) != dn(main): d["flags"].append(f"BAD:Ду не совпадает (ТЗ {dn(nm)}, счёт {dn(main)})")
            if pn(nm) and pn(main) and pn(main) < pn(nm): d["flags"].append(f"BAD:давление ниже ТЗ (Ру{pn(nm):g} → Ру{pn(main):g})")
            if re.search(r"электропривод|ЭИМ|с эл\.", nm, re.I) and not any(re.search(r"привод|ЭИМ|МЭО|GAIR", x["name"]) for x in lines): d["flags"].append("COND:в ТЗ с электроприводом — подтвердить, что привод входит в цену")
            if CN.search(main): d["flags"].append("COND:КНР/импорт — номер в реестре Минпромторга")
            if "КОФ" in nm.upper() and len(lines) == 1: d["flags"].append("COND:в ТЗ с КОФ — докупить КОФ")
            if "30нж" in nm and re.search(r"CF8|304", main) and "12Х18Н9ТЛ" in tz["Материал корпуса"].upper().replace(" ", ""): d["flags"].append("COND:корпус CF8/304 вместо 12Х18Н9ТЛ — согласие заказчика")
        elif d["tt"] == "1.1/24":
            d.update(inv=TULA["items"][0]["unit_vat"], sup="Тулаэлектропривод", offer=TULA["items"][0]["name"], srok="75–90 дн., 100% предоплата")
        elif d["tt"] in ADL:
            d.update(inv=float(ADL[d["tt"]]["unit_vat"]), sup="ТД АДЛ (Гранвэл)", offer=ADL[d["tt"]]["name"] + " — Россия, ЗПВС в реестре МПТ", srok="0–2 нед., 100% предоплата, самовывоз Лыткарино")
        elif d["tt"] in LIDER:
            d.update(inv=float(LIDER[d["tt"]]["unit_vat"]), sup="Лидер-Энерго", offer=LIDER[d["tt"]]["name"] + " (подтвердить полное обозначение по ТЗ)", srok="1–3 раб. дня, 100% предоплата")
        elif d["tt"] in KIP:
            d.update(inv=float(KIP[d["tt"]]["unit_vat"]), sup="КИПИКА (прайс 01.10)", offer=KIP[d["tt"]]["name"], srok="со склада, доставка ТК")
            d["flags"].append("COND:" + KIP[d["tt"]]["check"])
        d["budget"] = int(d["our"] * (1 - TORG) // 100 * 100)
        if d["inv"] is not None:
            dev = d["inv"] / d["our"] - 1; d["dev"] = dev
            bad = [f[4:] for f in d["flags"] if f.startswith("BAD")]; cond = [f[5:] for f in d["flags"] if f.startswith("COND")]
            if bad: d.update(act="Заказать у других", why="; ".join(bad))
            elif dev > 0.30: d.update(act="Заказать у других", why=f"дороже нашей цены на {dev:+.0%}")
            elif dev > 0.10: d.update(act="Торг: наш бюджет", why=f"дороже на {dev:+.0%}")
            elif cond: d.update(act="Брать при условии", why="; ".join(cond))
            else: d.update(act="Брать", why="в пределах нашей цены (±10%)" if dev >= -0.1 else f"дешевле нашей на {dev:+.0%}")
        else:
            d["dev"] = None; p = PEND.get(g, {}).get(z)
            if p: d.update(act="Не отработано", why=p)
            elif fact(d["tt"]): d.update(act="Не отработано", why="запрос КП проверенным заводам ещё не отправлен / ответа нет")
            else: d.update(act="Не отработано", why="ДН.РУ не котировал — запросить у заводов")
        d["alt"] = (fact(d["tt"]) or alt(d["name"])) if d["act"] != "Брать" else ""
        if d["sup"] and d["act"] != "Брать": d["alt"] = "; ".join(x for x in d["alt"].split("; ") if d["sup"] not in x)
        d["vs_old"] = (d["inv"] / d["old"] - 1) if d.get("old") and d["inv"] else None
        rows.append(d)

ORDER = ["Брать", "Брать при условии", "Торг: наш бюджет", "Заказать у других", "Не отработано"]
COL = {"Брать": "C6EFCE", "Брать при условии": "E2EFDA", "Торг: наш бюджет": "FFEB9C", "Заказать у других": "F8CBAD", "Не отработано": "DDEBF7"}
WHAT = {"Брать": "заказывать по счёту; по всем позициям DN.ru запросить страну производства и номер в реестре МПТ (ПП 1875); счёт ДН.РУ действует 2 дня",
        "Брать при условии": "брать после письменного подтверждения: реестр МПТ / эквивалент / КОФ / корпус CF8",
        "Торг: наш бюджет": "назвать наш бюджет; не соглашаются — к заводам из колонки «Кому ещё»",
        "Заказать у других": "у этого поставщика не брать: дороже рынка > 30% или не по ТЗ",
        "Не отработано": "цены нет: ДН.РУ подбирает / ждёт ответа, или запрос заводам ещё не отправлен"}
wb = openpyxl.Workbook(); sv = wb.active; sv.title = "Итог"
sv.append(["Камчатскэнерго 32616346943 — статус закупки на 26.10.2026: что берём, где торгуемся, что заказываем у других, что не отработано"]); sv["A1"].font = Font(bold=True, size=12)
sv.append(["Источники: счета ДН.РУ 819892 (52/53) и 819896 (96/109) от 26.10; КП Тулаэлектропривод 10/160; КП Лидер-Энерго 2022, КП АДЛ УТД000112759 (Гранвэл), прайс КИПИКА 01.10.2026 (краны для манометров). Бюджет = наша проверенная цена − 10% (себестоимость в расчёте лота). Цены с НДС."])
sv.append(["Решение", "Позиций", "Шт", "Наша цена × кол-во, ₽", "Цена поставщика × кол-во, ₽", "Бюджет × кол-во, ₽", "Что делать"])
for c in sv[3]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")
for a in ORDER:
    g = [d for d in rows if d["act"] == a]
    sv.append([a, len(g), sum(d["q"] for d in g), sum(d["our"] * d["q"] for d in g), sum((d["inv"] or 0) * d["q"] for d in g) or None, sum(d["budget"] * d["q"] for d in g), WHAT[a]])
    for j in (4, 5, 6): sv.cell(sv.max_row, j).number_format = M
    sv.cell(sv.max_row, 1).fill = PatternFill("solid", fgColor=COL[a])
sv.append(["ИТОГО", len(rows), sum(d["q"] for d in rows), sum(d["our"] * d["q"] for d in rows), None, sum(d["budget"] * d["q"] for d in rows), ""])
for j in (4, 6): sv.cell(sv.max_row, j).number_format = M
for c in sv[sv.max_row]: c.font = Font(bold=True)
sv.append([])
q = [d for d in rows if d["inv"] is not None and d["act"] != "Заказать у других"]
so, si = sum(d["our"] * d["q"] for d in q), sum(d["inv"] * d["q"] for d in q)
cb = sum(min(d["inv"], d["our"] * (1 - TORG)) * d["q"] for d in q)
sv.append(["Котировано и годно (брать + при условии + торг)", len(q), "", so, si, None, f"поставщик {si / so - 1:+.1%} к нашей цене"])
for j in (4, 5): sv.cell(sv.max_row, j).number_format = M
chg = [d for d in rows if d["vs_old"] is not None and abs(d["vs_old"]) > 0.005]
sv.append(["ДН.РУ: позиций с изменённой ценой против прошлого счёта", len(chg), "", None, None, None, "; ".join(f'{d["grp"][0]}{d["z"]}: {d["vs_old"]:+.0%}' for d in chg)[:250] or "цены не менялись"])
new = [d for d in rows if d["sup"] == "ДН.РУ" and not d.get("old")]
sv.append(["ДН.РУ: новые позиции в счёте (не было в прошлом)", len(new), "", None, None, None, "; ".join(f'{d["grp"]} №{d["z"]}' for d in new)[:250]])
sv.append([])
sv.append(["Ответ ДН.РУ по Kvs (письмо 26.10) — из ТЗ, Прил. 1, табл. 1.3", "", "", "", "", "", ""])
for c in sv[sv.max_row]: c.font = Font(bold=True)
for t in ["поз. 94 (25ч945п Ду15, табл. 1.3 №2): пропускная способность — 4 м³/ч",
          "поз. 95 (25ч945п Ду25, табл. 1.3 №3): пропускная способность — 4 м³/ч",
          "поз. 96 (25ч945п Ду50 «KV100», табл. 1.3 №4): максимальная пропускная способность — 65 м³/ч, допустимо 40 м³/ч (KV100 в наименовании — ошибка ТЗ)",
          "поз. 97 (25ч945п Ду80, табл. 1.3 №5): в ТЗ указано 4 м³/ч — для Ду80 нетипично, уточнить у заказчика через запрос разъяснений; ДН.РУ предварительно дать 4 м³/ч"]:
    sv.append([t]); sv.merge_cells(start_row=sv.max_row, start_column=1, end_row=sv.max_row, end_column=7)
for row in sv.iter_rows(min_row=3):
    for c in row:
        if c.value is not None: c.border = BOX
for i, w in enumerate([46, 9, 8, 18, 18, 18, 80], 1): sv.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

H = ["Решение", "Группа", "№ в заявке", "Таблица/№ ТЗ", "Позиция ТЗ", "Кол-во", "Наша цена, ₽/шт", "Цена поставщика, ₽/шт", "Отклонение", "Было в прошлом счёте, ₽/шт", "Бюджет для торга, ₽/шт",
     "Наша сумма, ₽", "Сумма поставщика, ₽", "Поставщик", "Что предложено", "Срок", "Почему", "Кому ещё / заводы"]


def sheet(title, data):
    ws = wb.create_sheet(title); ws.append(H)
    for c in ws[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")
    for d in data:
        ws.append([d["act"], d["grp"], d["z"], d["tt"], d["name"][:90], d["q"], d["our"], d["inv"], d["dev"], d.get("old"), d["budget"], d["our"] * d["q"],
                   d["inv"] * d["q"] if d["inv"] else None, d["sup"], d["offer"][:120], d["srok"], d["why"], d["alt"][:250]])
        X = ws.max_row
        for j in (7, 8, 10, 11, 12, 13): ws.cell(X, j).number_format = M
        ws.cell(X, 9).number_format = P; ws.cell(X, 1).fill = PatternFill("solid", fgColor=COL[d["act"]])
        if d["dev"] is not None: ws.cell(X, 9).fill = PatternFill("solid", fgColor="C6EFCE" if d["dev"] <= 0.1 else "FFEB9C" if d["dev"] <= 0.3 else "F8CBAD")
        for c in ws[X]: c.border = BOX
    for i, w in enumerate([18, 11, 7, 9, 38, 7, 12, 12, 9, 12, 12, 13, 13, 14, 36, 13, 36, 50], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = "F2"; ws.auto_filter.ref = f"A1:R{ws.max_row}"


rows.sort(key=lambda d: (ORDER.index(d["act"]), ["Задвижки", "Клапаны", "Краны", "Затворы", "Приводы и прочее"].index(d["grp"]), -(d["our"] * d["q"])))
sheet("Все позиции (315)", rows)
for a in ORDER: sheet({"Торг: наш бюджет": "Торг (бюджет)"}.get(a, a), [d for d in rows if d["act"] == a])
tidy(wb, row_h=40)
wb.save("../KAMCHATKA_STATUS_2610.xlsx")
for row in sv.iter_rows(min_row=4, max_row=13, values_only=True): print([x if not isinstance(x, float) else round(x) for x in row[:7]])
