"""Камчатскэнерго 32616346943: сверка счёта ДН.РУ № 819896 (клапаны, лист «Клапаны») и КП Тулаэлектропривода № 10/160 (поз. 1.1/24) с нашей проверенной ценой."""
import json, re, openpyxl
from collections import defaultdict
from openpyxl.styles import Font, PatternFill, Border, Side
from tidy import tidy

INV = json.load(open("data/kamch/kp/dnru_819896.json")); TULA = json.load(open("data/kamch/kp/tula_10-160.json"))
G = defaultdict(list)
for x in INV["items"]: G[x["z"]].append(x)
wbk = openpyxl.load_workbook("../KAMCHATKA_BY_GROUP_KP.xlsx")
POS = [r for r in wbk["Клапаны"].iter_rows(values_only=True) if r[0] and str(r[0]).isdigit() and re.match(r"^\d\.\d+ / \d+", str(r[1] or ""))]
ACT = [r for r in wbk["Приводы и прочее"].iter_rows(values_only=True) if r[0] and str(r[0]).isdigit() and re.match(r"^\d\.\d+ / \d+", str(r[1] or ""))]
TORG = 0.10; COST_KAM = 54810725
thin = Side(style="thin", color="BFBFBF"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin); M = '#,##0" ₽"'; P = "+0%;-0%;0%"
F = {"g": "C6EFCE", "y": "FFEB9C", "r": "F8CBAD", "b": "DDEBF7"}
MARK = re.compile(r"\b(\d{2}[а-яa-z]{1,2}\d{1,3}[а-яa-z]{0,3}\d{0,2}|VFG-2R|VFM|AFD|NVD|РД-НО|УРРД|ПСК|ВДХ|АРТ)", re.I)
CN = re.compile(r"Китай|BasicStandard|Benarmo|Dendor|Мегастрой", re.I)


def head(ws):
    for c in ws[ws.max_row]:
        if c.value is not None: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")


def pn(s):
    m = re.search(r"(?:Ру|PN)\s?(\d+(?:[.,]\d+)?)", s); return float(m.group(1).replace(",", ".")) if m else None


def dn(s):
    m = re.search(r"(?:Ду|DN|Д)\s?(\d+)", s); return m.group(1) if m else None


rows = []
for r in POS:
    z, tt, nm, q, our = int(r[0]), r[1], str(r[2]).split("·", 1)[-1].strip(), float(r[7]), float(r[8])
    sub = str(r[2]).split("·")[0].strip()
    d = dict(z=z, tt=tt, sub=sub, name=nm, q=q, our=our, inv=None, offer="", srok="", note="", col="b")
    lines = G.get(z)
    if not lines:
        d["note"] = "Нет в счёте — ДН.РУ не котировал"; rows.append(d); continue
    main = lines[0]; inv = sum(x["sum_vat"] for x in lines)
    d.update(inv=inv, dev=inv / our - 1, offer=main["name"] + (f' + КОФ ({len(lines) - 1} строк: фланцы, прокладки, крепёж)' if len(lines) > 1 else ""), srok=main["srok"])
    notes = []
    mk = [m.lower() for m in MARK.findall(nm)]
    if mk and not any(m in main["name"].lower() for m in mk): notes.append(("ДРУГОЙ ТОВАР (задвижка вместо клапана)" if "Задвижка" in main["name"] and "задвижк" not in nm.lower() else f"АНАЛОГ, не марка ТЗ ({mk[0]}) — нужно обоснование эквивалентности") )
    if dn(nm) and dn(main["name"]) and dn(nm) != dn(main["name"]): notes.append(f"Ду не совпадает: ТЗ {dn(nm)}, счёт {dn(main['name'])}")
    if pn(nm) and pn(main["name"]) and pn(main["name"]) < pn(nm): notes.append(f"Ду ок, но давление ниже ТЗ: ТЗ Ру{pn(nm):g}, счёт Ру{pn(main['name']):g}")
    if "КОФ" in nm.upper() and len(lines) == 1: notes.append("в ТЗ «с КОФ» — в счёте без ответных фланцев")
    if CN.search(main["name"]): notes.append("производство КНР / импорт — по ПП 1875 (ограничение) не пройдёт без реестровой записи")
    if d["dev"] > 0.3: notes.append("дороже нашей цены > 30%")
    elif d["dev"] > 0.1: notes.append("дороже нашей цены — торговаться")
    elif d["dev"] < -0.3: notes.append("дешевле нашей цены > 30% — проверить соответствие")
    elif d["dev"] < -0.1: notes.append("дешевле нашей цены")
    else: notes.append("в пределах ±10%")
    bad = any(n.startswith(("ДРУГОЙ", "Ду не", "Ду ок, но давление")) for n in notes); d["analog"] = any(n.startswith("АНАЛОГ") for n in notes)
    d["col"] = "r" if bad or d["dev"] > 0.3 else "y" if len(notes) > 1 or d["dev"] > 0.1 else "g"
    d["bad"] = bad; d["cn"] = bool(CN.search(main["name"])); d["note"] = "; ".join(notes)
    rows.append(d)

have = [d for d in rows if d["inv"] is not None]; ok = [d for d in have if not d["bad"]]
s_our = sum(d["our"] * d["q"] for d in ok); s_inv = sum(d["inv"] * d["q"] for d in ok)
c_our = s_our * (1 - TORG); c_best = sum(min(d["inv"], d["our"] * (1 - TORG)) * d["q"] for d in ok)
c_best_ru = sum((min(d["inv"], d["our"] * (1 - TORG)) if not d["cn"] else d["our"] * (1 - TORG)) * d["q"] for d in ok)
a = next(r for r in ACT if "Н-В-02" in str(r[2])); tq, tour = float(a[7]), float(a[8]); tinv = TULA["items"][0]["unit_vat"]

wb = openpyxl.Workbook(); sv = wb.active; sv.title = "Итог"
sv.append(["Камчатскэнерго 32616346943 — сверка счетов поставщиков с нашей таблицей (23.10.2026)"]); sv["A1"].font = Font(bold=True, size=12)
sv.append(["Цены с НДС. Наша цена — проверенная рыночная (KAMCHATKA_BY_GROUP_KP). Себестоимость в расчёте лота = наша цена − 10% торг."])
sv.append(["Показатель", "Значение", "Комментарий"]); head(sv)
for a_, b_, c_, f_ in [
    ("Счёт ДН.РУ № 819896 (клапаны): строк / сумма", f'{len(INV["items"])} строк', f'{INV["total_vat"]:,.2f} ₽, цены за 1 шт, действует 2 дня'.replace(",", " "), None),
    ("Позиций листа «Клапаны»", len(rows), f"{int(sum(d['q'] for d in rows))} шт", "0"),
    ("Котировано ДН.РУ", len(have), "", "0"),
    ("— из них другой товар / другой Ду (не по ТЗ)", len(have) - len(ok), "в сравнение не включены", "0"),
    ("— из них аналоги (не марка из ТЗ)", sum(1 for d in ok if d["analog"]), "лот допускает эквиваленты с обоснованием — проверить характеристики", "0"),
    ("— из них производство КНР / импорт", sum(1 for d in ok if d["cn"]), "по ПП 1875 (ограничение) — риск отклонения", "0"),
    ("Не котировано", len(rows) - len(have), "запросить у заводов (Теплоконтроль, Армагус, Ридан и др.)", "0"),
    ("Сопоставимые позиции: наша цена × кол-во", s_our, "", M),
    ("Сопоставимые позиции: счёт × кол-во", s_inv, f"{s_inv / s_our - 1:+.0%} к нашей цене", M),
    ("Себестоимость в расчёте (−10%)", c_our, "", M),
    ("Если брать у ДН.РУ только где дешевле", c_best, f"{c_best - c_our:+,.0f} ₽ к расчёту".replace(",", " "), M),
    ("То же, без китайских / импортных", c_best_ru, f"{c_best_ru - c_our:+,.0f} ₽ к расчёту — реалистично".replace(",", " "), M),
]:
    sv.append([a_, b_, c_]);
    if f_: sv.cell(sv.max_row, 2).number_format = f_
sv.append([])
sv.append(["Тулаэлектропривод КП № 10/160 от 23.10.2026", "", ""]); head(sv)
for a_, b_, c_, f_ in [("Позиция", "1.1/24 Электропривод Н-В-02 У2 (под задвижку 30с964нж Ду300)", "", None),
                       ("Цена КП за шт с НДС", tinv, "167 200 без НДС", M), ("Наша цена за шт", tour, f"КП дороже на {tinv / tour - 1:+.0%}", M),
                       ("Кол-во по ТЗ", tq, "КП выставлено на 1 шт — нужно 2", "0"), ("Сумма по КП на 2 шт", tinv * tq, f"{(tinv - tour * (1 - TORG)) * tq:+,.0f} ₽ к себестоимости в расчёте".replace(",", " "), M),
                       ("Условия", TULA["terms"], "срок 75–90 дн. укладывается в 120 дн. поставки; 100% предоплата — наши деньги на 3 мес.", None),
                       ("Плюс", "российский завод (Тула) — проходит нацрежим; номер в реестре Минпромторга запросить", "", None)]:
    sv.append([a_, b_, c_])
    if f_: sv.cell(sv.max_row, 2).number_format = f_
sv.append([])
d_ = (c_best_ru - c_our) + (tinv - tour * (1 - TORG)) * tq
sv.append(["Влияние на себестоимость лота Камчатки", COST_KAM + d_, f"было {COST_KAM:,.0f} ₽; {d_:+,.0f} ₽ (клапаны у ДН.РУ где дешевле и не КНР + приводы Тула)".replace(",", " ")])
sv.cell(sv.max_row, 2).number_format = M
sv.append([])
sv.append(["По подгруппам (сопоставимые)", "Наша × кол-во / Счёт × кол-во", "Отклонение"]); head(sv)
for sg in dict.fromkeys(d["sub"] for d in ok):
    gg = [d for d in ok if d["sub"] == sg]; a1, b1 = sum(d["our"] * d["q"] for d in gg), sum(d["inv"] * d["q"] for d in gg)
    sv.append([f"{sg} ({len(gg)} поз.)", f"{a1:,.0f} / {b1:,.0f} ₽".replace(",", " "), b1 / a1 - 1]); sv.cell(sv.max_row, 3).number_format = P
for row in sv.iter_rows(min_row=3):
    for c in row:
        if c.value is not None: c.border = BOX
for i, w in enumerate([48, 40, 70], 1): sv.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

ws = wb.create_sheet("Клапаны построчно")
ws.append(["№ по заявке", "ТТ / №", "Подгруппа", "Позиция ТЗ", "Кол-во", "Наша цена, ₽/шт", "Счёт ДН.РУ, ₽/шт (с КОФ)", "Отклонение", "Наша сумма, ₽", "Сумма по счёту, ₽", "Предложено в счёте", "Срок", "Вывод"]); head(ws)
for d in rows:
    ws.append([d["z"], d["tt"], d["sub"], d["name"][:90], d["q"], d["our"], d["inv"], d.get("dev"), d["our"] * d["q"], d["inv"] * d["q"] if d["inv"] else None, d["offer"][:160], d["srok"], d["note"]])
    X = ws.max_row
    for j in (6, 7, 9, 10): ws.cell(X, j).number_format = M
    ws.cell(X, 8).number_format = P
    for j in (7, 8, 13): ws.cell(X, j).fill = PatternFill("solid", fgColor=F[d["col"]])
    for c in ws[X]: c.border = BOX
for i, w in enumerate([8, 9, 18, 40, 7, 13, 13, 10, 14, 14, 50, 14, 50], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "E2"; ws.auto_filter.ref = f"A1:M{ws.max_row}"
tidy(wb, row_h=44)
wb.save("../KAMCHATKA_KP_DNRU_819896_TULA.xlsx")
print(len(have), "котир.;", len(have) - len(ok), "не по ТЗ;", sum(1 for d in ok if d["cn"]), "КНР | наша", round(s_our), "счёт", round(s_inv), f"{s_inv / s_our - 1:+.1%}",
      "| себес", round(c_our), "лучшее", round(c_best), "без КНР", round(c_best_ru), "| тула", tinv, tour, tq, "| лот", round(COST_KAM + d_))
for d in rows:
    if d.get("bad"): print("BAD", d["z"], d["name"][:50], "|", d["offer"][:60])
