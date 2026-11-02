"""Камчатка 32616346943: у кого выгодно брать. Все полученные КП/счета построчно против нашей цены в заявке (84 518 812,74 ₽ без НДС, мы на УСН —
закупка с НДС = себестоимость). Доставка до Камчатки 3 069 024 ₽ распределена пропорционально выручке позиции.
Выгодно = маржа после доставки > 0. Соответствие ТЗ — по построчной проверке (data/kamch/check). → ../KAMCHATKA_VYGODNO_0211.xlsx"""
import json, re, io, contextlib, runpy, openpyxl
from collections import defaultdict
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from tidy import tidy

with contextlib.redirect_stdout(io.StringIO()):
    S = runpy.run_path("kamch_status.py")
ROWS = {d["tt"]: d for d in S["rows"]}; INV = S["INV"]
TZ = {f'{t["table"]}/{t["n_table"]}': t for t in json.load(open("data/kamch/tz_verbatim.json"))}
PR = json.load(open("data/kamch/forms/kp_prices_84518812.json")); FM = json.load(open("data/kamch/forms/kp_form_map.json"))
REV = sum(PR[t] * float(TZ[t]["qty"]) for t in PR); FREIGHT = 3069024.0
kp = lambda f: json.load(open(f"data/kamch/kp/{f}.json"))
CHK = {}
for f in ("data/kamch/check/out_dnru.json", "data/kamch/check/out_other.json"):
    for x in json.load(open(f)): CHK[(x["tt"], x.get("supplier") or "ДН.РУ")] = x

O = defaultdict(list)  # tt -> [(поставщик, предложение, цена ед. с НДС, условия)]
for d in ROWS.values():  # ДН.РУ (счета 819892/819896 v2, строка + КОФ)
    lines = INV.get(d["grp"], {}).get(d["z"])
    if lines: O[d["tt"]].append(("ДН.РУ", lines[0]["name"] + (" + КОФ" if len(lines) > 1 else ""), sum(x["sum_vat"] for x in lines), "склад Москва, 100% предоплата"))
for f, nm, terms in (("adl_utd112759", "ТД АДЛ", "0–2 нед., самовывоз Лыткарино"), ("lider-energo_2022", "Лидер-Энерго", "1–3 раб. дня"), ("kipika_0110", "КИПИКА", "со склада"),
                     ("epas_2710", "ЭПАС", "")):
    d0 = kp(f)
    for x in d0["items"]: O[x["tt"]].append((nm, x["name"], float(x["unit_vat"]), terms or str(d0.get("terms", ""))[:60]))
for x in kp("tula_10-160")["items"]: O[x["z"]].append(("Тулаэлектропривод", x["name"], float(x["unit_vat"]), "75–90 дн."))
for f, nm in (("teplogrand_match", "Теплогранд (LD) вар.1"), ("teplogrand_v2_match", "Теплогранд (LD) вар.2")):
    d0 = kp(f)
    for x in d0["out"] + [dict(v, tt=k) for k, v in d0.get("pdz", {}).items()]:
        if x.get("bad"): continue  # латунь вместо стали и т.п. — не по ТЗ
        O[x["tt"]].append((nm, x["offer"] + (" + КОФ" if x.get("kof") else ""), float(x["price"] + x.get("kof", 0)), "склад, самовывоз"))
B = kp("baz_00077")
for n, t, nm, p in B["items"]: O[t].append(("БАЗ", nm, round(p * 1.22, 2), "75–90 дн., самовывоз Благовещенск (РБ)"))

def verdict(t, sup):
    k = "ДН.РУ" if sup == "ДН.РУ" else "Теплогранд (LD)" if sup.startswith("Теплогранд") else {"ТД АДЛ": "ТД АДЛ (Гранвэл)", "КИПИКА": "КИПИКА (прайс 01.10)"}.get(sup, sup)
    v = CHK.get((t, k), {}).get("verdict_class") or str(CHK.get((t, k), {}).get("verdict") or "")
    return "не соответствует" if v.startswith("не соотв") else "при условии" if v.startswith("соотв") else "не проверено"

BRAND = [(r"dendor", "7805499441"), (r"УРРД", "0323085570"), (r"Гранлок|CV16", "7718625072"), (r"ЗПД|ЗДП", "6324029755"), (r"Гранвэл|ГРАНВЭЛ", "7718625072")]
MK = {"Теплогранд": ("7451211335", "7449114775"), "ТД АДЛ": ("7718625072",), "БАЗ": ("0258001489",), "Тулаэлектропривод": ("7130025182",), "Лидер-Энерго": ("6324029755",)}  # ИНН изготовителей, чей товар поставляет поставщик
order = sorted(TZ, key=lambda t: (t.split("/")[0], int(t.split("/")[1])))
res = []
for t in order:
    q = float(TZ[t]["qty"]); p = PR[t]; fr = FREIGHT * p / REV  # доставка на ед.
    offs = []
    for sup, name, c, terms in O.get(t, []):
        m = p - c - fr; offs.append(dict(sup=sup, name=name, c=c, terms=terms, m=m, mp=m / p, v=verdict(t, sup)))
    man = (FM[t]["val"] or ["", "", ""])[2]
    for x in offs:
        x["mok"] = any(i in man for i in MK.get(x["sup"].split(" (")[0].split(" вар")[0], ())) or \
            (x["sup"] == "ДН.РУ" and any(re.search(b, x["name"], re.I) and i in man for b, i in BRAND))  # ДН.РУ перепродаёт товар российских заводов (Dendor → АО «Энергия» и т.д.)
        x["why"] = "; ".join(w for w in ("" if x["mok"] else "производитель ≠ заявленному в КП (" + (man.split(" (ИНН")[0] or "—") + ")",
                                          "не соответствует ТЗ" if x["v"] == "не соответствует" else "") if w)
    offs.sort(key=lambda x: -x["m"])
    for x in offs:  # товар завода с записью ПП 719 (БАЗ, LD, АДЛ, Тула, Транс-Энерго), но в КП заявлен другой завод — исправимо: сменить производителя в КП
        x["fix"] = (not x["mok"]) and x["sup"].split(" (")[0].split(" вар")[0] in MK
        if x["fix"]: x["why"] = x["why"].replace("производитель ≠ заявленному в КП", "СМЕНИТЬ в КП производителя на завод поставщика; сейчас заявлен")
    good = [x for x in offs if (x["mok"] or x["fix"]) and x["v"] != "не соответствует" and x["m"] > 0]
    prof = [x for x in offs if x["m"] > 0]
    best = (good or prof or offs or [None])[0]
    cat = "БРАТЬ" if good else "ВЫГОДНО, НО УСЛОВИЕ" if prof else "НЕВЫГОДНО" if offs else "НЕТ КП"
    res.append(dict(t=t, q=q, p=p, fr=fr, offs=offs, best=best, cat=cat, name=TZ[t]["Наименование"], man=man))

thin = Side(style="thin", color="BFBFBF"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin); M = "#,##0"
HD = lambda ws: [setattr(c, "font", Font(bold=True, color="FFFFFF")) or setattr(c, "fill", PatternFill("solid", fgColor="305496")) or setattr(c, "alignment", Alignment(wrap_text=True)) for c in ws[ws.max_row]]
G, Y, R = (PatternFill("solid", fgColor=c) for c in ("E2EFDA", "FFF2CC", "F8CBAD"))
wb = openpyxl.Workbook(); sv = wb.active; sv.title = "Итог по поставщикам"
win = [r for r in res if r["cat"] == "БРАТЬ"]; cond = [r for r in res if r["cat"] == "ВЫГОДНО, НО УСЛОВИЕ"]
loss = [r for r in res if r["cat"] == "НЕВЫГОДНО"]; none = [r for r in res if r["cat"] == "НЕТ КП"]
sv.append([f"Камчатка 32616346943 — у кого выгодно брать (все КП на 02.11.2026). Цена заявки 84 518 812,74 ₽ без НДС; закупка с НДС; доставка {FREIGHT:,.0f} ₽ распределена по выручке"]); sv["A1"].font = Font(bold=True, size=12)
sv.append(["БРАТЬ: маржа > 0, производитель = заявленному в КП (реестр ПП 719), не «не соответствует» ТЗ"]); sv.cell(2, 1).font = Font(bold=True)
sv.append(["Поставщик", "Позиций", "Закупка с НДС, ₽", "Выручка, ₽", "Доставка, ₽", "Маржа, ₽", "Маржа, %"]); HD(sv)
def block(L):
    agg = defaultdict(lambda: [0, 0, 0, 0, 0])
    for r in L:
        b = r["best"]; a = agg[b["sup"]]; a[0] += 1; a[1] += b["c"] * r["q"]; a[2] += r["p"] * r["q"]; a[3] += r["fr"] * r["q"]; a[4] += b["m"] * r["q"]
    for k, a in sorted(agg.items(), key=lambda x: -x[1][4]): sv.append([k, a[0], a[1], a[2], a[3], a[4], a[4] / a[2]])
    t_ = [sum(a[i] for a in agg.values()) for i in range(5)]
    sv.append(["ИТОГО", t_[0], t_[1], t_[2], t_[3], t_[4], t_[4] / t_[2] if t_[2] else 0]); [setattr(c, "font", Font(bold=True)) for c in sv[sv.max_row]]
    return agg, t_
agg, tot = block(win)
sv.append([]); sv.append(["ВЫГОДНО ПО ЦЕНЕ, НО: производитель не тот, что в КП (ДН.РУ — DN.ru/КНР, ЭПАС, КИПИКА), или не соответствует ТЗ — брать только после замены/подтверждения"]); sv.cell(sv.max_row, 1).font = Font(bold=True)
sv.append(["Поставщик", "Позиций", "Закупка с НДС, ₽", "Выручка, ₽", "Доставка, ₽", "Маржа, ₽", "Маржа, %"]); HD(sv)
agg2, tot2 = block(cond)
sv.append([]); sv.append(["Невыгодно (все КП дороже нашей цены)", len(loss), sum(r["best"]["c"] * r["q"] for r in loss), sum(r["p"] * r["q"] for r in loss), "", sum(r["best"]["m"] * r["q"] for r in loss)])
sv.append(["Нет ни одного КП", len(none), "", sum(r["p"] * r["q"] for r in none)])
for row in sv.iter_rows(min_row=3):
    for j in (2, 3, 4, 5): row[j].number_format = M
    row[6].number_format = "0%"
for i, w in enumerate([34, 10, 16, 16, 13, 15, 9, 12, 12], 1): sv.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

ws = wb.create_sheet("Брать — по позициям")
ws.append(["Решение", "Поставщик", "Табл./№", "Позиция ТЗ", "Кол-во", "Предложение поставщика", "Цена ед. с НДС, ₽", "Сумма закупки, ₽", "Наша цена ед. (заявка), ₽", "Выручка, ₽", "Маржа после доставки, ₽", "Маржа, %",
           "Соответствие ТЗ", "Условия", "Альтернатива (2-е место)", "Производитель в нашем КП (реестр)", "Почему не «брать»"]); HD(ws)
for r in sorted(win + cond, key=lambda r: (r["cat"], r["best"]["sup"], r["t"])):
    b = r["best"]; alt = next((f'{x["sup"]}: {x["c"]:,.0f} ({x["v"]})' for x in r["offs"] if x is not b), "")
    ws.append([r["cat"], b["sup"], r["t"], r["name"], r["q"], b["name"], b["c"], b["c"] * r["q"], r["p"], r["p"] * r["q"], b["m"] * r["q"], b["mp"], b["v"], b["terms"], alt, r["man"], b["why"]])
    X = ws.max_row; ws.cell(X, 1).fill = G if r["cat"] == "БРАТЬ" else Y
    for j in (7, 8, 9, 10, 11): ws.cell(X, j).number_format = M
    ws.cell(X, 12).number_format = "0%"
for i, w in enumerate([14, 20, 8, 40, 7, 50, 13, 14, 13, 14, 14, 8, 15, 22, 30, 34, 40], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "E2"; ws.auto_filter.ref = f"A1:Q{ws.max_row}"

wl = wb.create_sheet("Невыгодно и нет КП")
wl.append(["Табл./№", "Позиция ТЗ", "Кол-во", "Наша цена ед., ₽", "Выручка, ₽", "Лучшее КП", "Цена ед. с НДС, ₽", "Убыток на позиции, ₽", "Нужна закупочная цена с НДС не выше, ₽"]); HD(wl)
for r in loss + none:
    b = r["best"]
    wl.append([r["t"], r["name"], r["q"], r["p"], r["p"] * r["q"], f'{b["sup"]}: {b["name"]}' if b else "КП нет", b["c"] if b else None, b["m"] * r["q"] if b else None, r["p"] - r["fr"]])
    X = wl.max_row; wl.cell(X, 6).fill = R if b else Y
    for j in (4, 5, 7, 8, 9): wl.cell(X, j).number_format = M
for i, w in enumerate([8, 44, 7, 13, 14, 50, 13, 15, 16], 1): wl.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
wl.freeze_panes = "C2"

wa = wb.create_sheet("Все КП по позициям")
wa.append(["Табл./№", "Позиция ТЗ", "Кол-во", "Наша цена ед., ₽", "Поставщик", "Предложение", "Цена ед. с НДС, ₽", "Маржа ед. после доставки, ₽", "Маржа, %", "Соответствие ТЗ", "Лучшее", "Замечание"]); HD(wa)
for r in res:
    for x in r["offs"]:
        wa.append([r["t"], r["name"], r["q"], r["p"], x["sup"], x["name"], x["c"], x["m"], x["mp"], x["v"], "★" if x is r["best"] else "", x["why"]])
        X = wa.max_row; wa.cell(X, 8).fill = G if x["m"] > 0 else R
        for j in (4, 7, 8): wa.cell(X, j).number_format = M
        wa.cell(X, 9).number_format = "0%"
for i, w in enumerate([8, 40, 7, 13, 20, 50, 13, 14, 8, 15, 6], 1): wa.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
wa.freeze_panes = "C2"; wa.auto_filter.ref = f"A1:L{wa.max_row}"
tidy(wb, row_h=30)
wb.save("../KAMCHATKA_VYGODNO_0211.xlsx")
for nm_, ag, tt_ in (("БРАТЬ", agg, tot), ("ВЫГОДНО, НО УСЛОВИЕ", agg2, tot2)):
    print("==", nm_)
    for k, a in sorted(ag.items(), key=lambda x: -x[1][4]): print(f"  {k:24} {a[0]:>4} закупка {a[1]:>13,.0f} выручка {a[2]:>13,.0f} маржа {a[4]:>12,.0f} ({a[4]/a[2]:.0%})")
    print("  ИТОГО", tt_[0], f"закупка {tt_[1]:,.0f} выручка {tt_[2]:,.0f} маржа {tt_[4]:,.0f}")
print("невыгодно", len(loss), f'выручка {sum(r["p"]*r["q"] for r in loss):,.0f}', "нет КП", len(none), f'выручка {sum(r["p"]*r["q"] for r in none):,.0f}')
