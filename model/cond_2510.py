"""Условия участия по лотам с подачей с 25.10 (без аукционов): аванс, обеспечение, оплата, сроки.
Источники: извещения ЕИС через gosplan (сырые JSON в каталоге argv[1]), выгрузки 19.10,
найденное на площадках/агрегаторах (argv[2], необязательно), нормы 44-ФЗ / ПП 1352."""
import json, re, sys, os, datetime as dt, openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from tidy import tidy

RAW = sys.argv[1]
WEB = {x["num"]: x for x in json.load(open(sys.argv[2]))} if len(sys.argv) > 2 and os.path.exists(sys.argv[2]) else {}
d = json.load(open("data/triage_1910.json"))
F = dt.datetime(2026, 10, 25)
NOT = re.compile(r"канцеляр|хозяйств|игрушк|реактив|лаб\. посуд|стоматолог|газон|офисн|учебн|театральн|сварк|мотопомп|инструмент|тара |укрывн|СМР|переустройств|реализаци|невостребован|запчаст|специнвентар|оснащени|пневматик|ЗИП|компьютерн|оргтехник|стройматериал", re.I)
ORDER = ["Соль", "Одежда/СИЗ", "Форма/снаряжение", "СИЗ (респираторы)", "Медрасходники (пробирки)", "Сантехника/арматура", "Кабель (ЭОМ)", "СС/видеонаблюдение", "Вентиляция/фильтры"]

sel, seen = [], set()
for x in d:
    if "аукцион" in (x["method"] or "").lower(): continue
    if re.search(r"работ|услуг|ремонт|монтаж", x["name"], re.I): continue
    try: e = dt.datetime.strptime(x["end"], "%Y-%m-%d %H:%M")
    except Exception: continue
    if e < F or NOT.search(x["name"]): continue
    k = (x["name"].strip().lower()[:60], (x["cust"] or "")[:30], x["nmck"])
    if k in seen: continue
    seen.add(k); x = dict(x); x["e"] = e; sel.append(x)
sel.sort(key=lambda x: (ORDER.index(x["cat"]), x["e"], -(x["nmck"] or 0)))


def load(n, k):
    f = f"{RAW}/{n}_{k}.json"
    if not os.path.exists(f): return None
    j = json.load(open(f))
    return None if "detail" in j else j


def num(v):
    try: return float(str(v).replace(" ", "").replace(",", "."))
    except Exception: return None


def short(t, n): return t if len(t) <= n else t[:n - 1].rstrip() + "…"


def rub(v): return f"{v:,.0f} ₽".replace(",", " ")


def deep(o, key):
    """Все значения ключа key на любой глубине."""
    out = []
    if isinstance(o, dict):
        for k, v in o.items():
            if k == key: out.append(v)
            out += deep(v, key)
    elif isinstance(o, list):
        for v in o: out += deep(v, key)
    return out


def from44(x):
    n = load(x["num"], "n")
    r = dict(app_rub=0, contr_rub=None)
    if not n:  # запрос цен — не закупка, извещения нет
        r.update(app="не требуется", contr="не требуется", adv="нет (это не закупка)",
                 pay="—", term="—", warr="—", pref="—",
                 note="Запрос цен — сбор КП для НМЦК, контракта нет. Подаём КП без обязательств.")
        return r
    cri = n["source"]["notificationInfo"]["customerRequirementsInfo"]["customerRequirementInfo"]
    if isinstance(cri, list): cri = cri[0]
    nmck = x["nmck"] or 0
    ag = cri.get("applicationGuarantee")
    if ag:
        a = num(ag.get("amount")) or nmck * num(ag.get("part")) / 100
        r["app"] = f"{rub(a)} ({num(ag.get('part')) or round(a / nmck * 100, 1)}%)"; r["app_rub"] = a
    else: r["app"] = "не требуется"
    cg = cri.get("contractGuarantee")
    if cg:
        p = num(cg.get("part")); a = num(cg.get("amount")) or (nmck * p / 100 if p else None)
        r["contr"] = f"{p or round(a / nmck * 100, 1)}% ≈ {rub(a)}" + (" (от цены контракта)" if not cg.get("amount") else "")
        r["contr_rub"] = a
    else: r["contr"] = "не требуется"
    w = cri.get("warrantyInfo") or {}
    wt = w.get("warrantyTerm") or w.get("manufacturersWarrantyRequirement") or ""
    m = re.search(r"не менее\s+(\d+)\s*(месяц|мес|год)", wt)
    r["warr"] = (f"не менее {m.group(1)} {m.group(2)}." if m else (short(wt, 40) if wt else "в извещении нет")) + \
        ("; обеспечение гарантии" if cri.get("guarantyObligationsInfo") or cri.get("provisionWarranty") else "")
    ex = deep(cri, "contractExecutionTermsInfo")
    t = ""
    if ex:
        nr, rl = ex[0].get("notRelativeTermsInfo"), ex[0].get("relativeTermsInfo")
        if nr: t = ("с " + nr["startDate"][:10] + " " if nr.get("startDate") else "") + "до " + nr.get("endDate", "")[:10]
        elif rl: t = f"{rl.get('term')} дн. с даты контракта"
    t = re.sub(r"(\d{4})-(\d\d)-(\d\d)", r"\3.\2.\1", t)
    r["term"] = t or "в извещении нет"
    pr = [p["preferenseRequirementInfo"]["shortName"] for p in deep(n, "preferenseInfo") for p in (p if isinstance(p, list) else [p])]
    pref = []
    if "PVS33044" in pr or "СМП" in (x["smp"] or ""): pref.append("только СМП/СОНО")
    if "IN44" in pr or "UG44" in pr: pref.append("преф. 15% инвалидам/УИС")
    if any(v == "true" for v in deep(n, "isRestrictForeignPurchaseObjects")): pref.append("нацрежим (ограничение иностр.)")
    r["pref"] = "; ".join(pref) or "—"
    r["adv"] = "не указан — сверить с проектом"
    r["pay"] = "≤ 7 раб. дн. после приёмки (ст. 34)"
    r["note"] = "ЗК: без торгов, побеждает мин. цена. Снижение ≥25% → обеспечение ×1,5." \
        if "котиров" in x["method"].lower() else ""
    return r


def from223(x):
    p = load(x["num"], "p")
    r = dict(app_rub=0, contr_rub=None)
    if not p:
        r.update(app="по документации", contr="по документации", adv="по документации", pay="по документации", term="—", warr="—", pref="—", note="Извещение не загрузилось")
        return r
    src = [dd["source"] for dd in p["docs"] if deep(dd["source"], "lotData")][-1]
    ld = deep(src, "lotData")[0]
    smp = ld.get("forSmallOrMiddle") == "true" or "МСП" in (src.get("purchaseCodeName") or "") or "СМП" in (x["smp"] or "")
    nmck = x["nmck"] or 0
    if ld.get("applicationSupplyNeeded") == "true":
        a = num(ld.get("applicationSupplySumm")) or num(x["sec_app"])
        r["app"] = f"{rub(a)} ({round(a / nmck * 100, 1)}%)" if a and nmck else "есть, сумма в документации"; r["app_rub"] = a or 0
    else: r["app"] = "не требуется"
    cs = ld.get("completingSupplyInfo") or {}
    if ld.get("completingSupplyNeeded") == "true":
        s = cs.get("sum") or ""
        a = num(s) or num(x["sec_contr"])
        if a and nmck: r["contr"] = f"≈ {rub(a)} ({round(a / nmck * 100, 1)}%)"
        elif "аванс" in s.lower(): r["contr"] = "на возврат аванса: 100% аванса (БГ)"
        else: r["contr"] = "есть, размер — в Прил. к документации"
        r["contr_rub"] = a
    elif num(x["sec_contr"]):
        a = num(x["sec_contr"]); r["contr"] = f"{rub(a)} ({round(a / nmck * 100, 1)}%)"; r["contr_rub"] = a
    else: r["contr"] = "в извещении не требуется"
    r["adv"] = "есть (см. документацию)" if "Есть" in str(x["advance"]) or "аванс" in str(cs).lower() else "не указан — см. проект договора"
    r["pay"] = "≤ 7 раб. дн. после приёмки (ПП 1352)" if smp else "по договору (обычно 30–60 дн.)"
    r["term"] = "по документации"
    r["warr"] = "по документации"
    pref = ["только МСП" if smp else ""]
    li = deep(src, "lotItem")
    li = [i for v in li for i in (v if isinstance(v, list) else [v])]
    if any(i.get("prohibitionNationalRegime") == "true" for i in li): pref.append("нацрежим: запрет иностр.")
    elif any(i.get("constraintNationalRegime") == "true" for i in li): pref.append("нацрежим: ограничение иностр.")
    elif any(i.get("superiorityNationalRegime") == "true" for i in deep(src, "lotItems") if isinstance(i, dict)): pref.append("нацрежим: преимущество РФ")
    r["pref"] = "; ".join(p for p in pref if p) or "—"
    r["note"] = "Только МСП: обесп. заявки ≤2%, договора ≤5% (ПП 1352)." if smp else ""
    return r


def fromcom(x):
    r = dict(app_rub=0, contr_rub=None)
    a = num(x["sec_app"]); c = num(x["sec_contr"])
    r["app"] = rub(a) if a else "не указано (обычно нет)"; r["app_rub"] = a or 0
    r["contr"] = rub(c) if c else "не указано"; r["contr_rub"] = c
    r["adv"] = "есть (см. документацию)" if "Есть" in str(x["advance"]) else "не указан — запросить"
    r["pay"] = "запросить (обычно 30–90 дн.)"
    r["term"] = r["warr"] = "по документации"; r["pref"] = "—"
    r["note"] = "Коммерческая: условия в документации площадки, возможна переторжка." + \
        (" Нет НМЦК — подаём КП." if not x["nmck"] else "")
    if "Мониторинг" in x["method"] or x["method"] == "Запрос цен": r["note"] = "Мониторинг/запрос цен — сбор КП без обязательств."
    return r


rows = []
for x in sel:
    law = x["law"]
    r = from44(x) if law.startswith("44") else from223(x) if law.startswith("223") else fromcom(x)
    w = WEB.get(x["num"])
    if w and w.get("found"):
        for k, kk in [("advance", "adv"), ("payment", "pay"), ("delivery", "term"), ("sec_app", "app"), ("sec_contr", "contr"), ("warranty", "warr")]:
            if w.get(k): r[kk] = short(w[k], 42) + " [пл.]"
        r["note"] = "Подробно — лист «С площадок». " + r["note"]
        sc = w.get("sec_contr") or ""
        m = re.search(r"(\d+)%\s*НМЦ", sc); m2 = re.search(r"([\d ]+),\d\d\s*руб", sc)
        if m and x["nmck"]: r["contr_rub"] = x["nmck"] * int(m.group(1)) / 100
        elif m2: r["contr_rub"] = num(m2.group(1))
        r["src"] = w.get("source", "")
    if x["e"].year > 2026: r["note"] = ("Срок подачи в карточке " + x["e"].strftime("%d.%m.%Y") + " — проверить. " + r["note"]).strip()
    # деньги, которые замораживаем: заявка + контракт (обеспечение контракта считаем от НМЦК — верхняя оценка)
    r["freeze"] = (r["app_rub"] or 0) + (r["contr_rub"] or 0)
    rows.append((x, r))

thin = Side(style="thin", color="BFBFBF"); MONEY = '#,##0" ₽"'
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Условия участия"
ws.append(["Условия участия по лотам с подачей с 25.10.2026 — без аукционов (наш профиль, без дублей)"]); ws["A1"].font = Font(bold=True, size=13)
ws.append(["44-ФЗ — из извещений ЕИС; 223-ФЗ — из извещений и норм ПП 1352; коммерческие — из карточек выгрузки 19.10. «[площадка]» — найдено на странице площадки/агрегатора. Документацию ЕИС скачать не удалось — аванс и срок оплаты сверять с проектом контракта."])
H = ["№", "Категория", "Номер", "Предмет", "Заказчик", "Закон", "Способ", "НМЦК, ₽", "Подача до", "Обеспечение заявки", "Обеспечение контракта", "Аванс",
     "Срок оплаты", "Срок поставки / исполнения", "Гарантия на товар", "Ограничения / преференции", "Замораживаем денег, ₽", "Комментарий"]
ws.append(H)
for c in ws[3]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")
for i, (x, r) in enumerate(rows, 1):
    ws.append([i, x["cat"], x["num"], short(x["name"], 105), short(x["cust"] or "", 75), x["law"], x["method"], x["nmck"], x["e"], r["app"], r["contr"], r["adv"], r["pay"], r["term"], r["warr"], r["pref"],
               r["freeze"] or None, r["note"]])
    R = ws.max_row; ws.cell(R, 3).number_format = "@"; ws.cell(R, 8).number_format = MONEY; ws.cell(R, 17).number_format = MONEY
    ws.cell(R, 9).number_format = "DD.MM.YYYY HH:MM"
    if "есть" in str(r["adv"]).lower() and "нет" not in str(r["adv"]).lower(): ws.cell(R, 12).fill = PatternFill("solid", fgColor="C6EFCE")
    if r["freeze"] and r["freeze"] > 500000: ws.cell(R, 17).fill = PatternFill("solid", fgColor="FCE4D6")
    for cc in ws[R]: cc.border = Border(top=thin, bottom=thin, left=thin, right=thin)
for i, w in enumerate([4, 15, 21, 44, 34, 11, 20, 14, 15, 18, 22, 20, 22, 18, 16, 24, 14, 40], 1):
    ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "E4"; ws.auto_filter.ref = f"A3:R{ws.max_row}"

lg = wb.create_sheet("Нормы — справка")
lg.append(["Что по закону, если в документации не сказано иначе"]); lg["A1"].font = Font(bold=True, size=13)
lg.append(["Тип", "Обеспечение заявки", "Обеспечение контракта", "Аванс", "Срок оплаты", "Важно"])
for c in lg[2]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")
for r in [
    ["44-ФЗ, запрос котировок", "не требуется (ч. 1 ст. 44 — только конкурс/аукцион)", "0,5–30% от цены контракта, если установлено; можно банковской гарантией или деньгами", "только если прописан в проекте контракта (до 30%)", "≤ 7 рабочих дней с подписания документа о приёмке", "Снижение ≥25% → обеспечение ×1,5 (ст. 37); опыт в 3 исполненных контракта даёт право на уменьшение обеспечения для СМП"],
    ["44-ФЗ, запрос цен", "нет", "нет", "нет", "—", "Не закупка: сбор КП для НМЦК, подача без обязательств"],
    ["223-ФЗ, только МСП (ПП 1352)", "≤ 2% НМЦК (или не требуется при НМЦК ≤ 20 млн)", "≤ 5% НМЦК или в размере аванса, если аванс > 5%", "по договору", "≤ 7 рабочих дней с подписания документа о приёмке", "Обеспечение можно банковской гарантией"],
    ["223-ФЗ, прочие", "по положению о закупке заказчика (часто 0,5–5%)", "по документации (обычно 5–30%)", "по договору", "по договору (обычно 30–60 дн.)", "Россети, ФСК и т. п.: часто аванс 0–30% при банковской гарантии на аванс"],
    ["Коммерческие", "как правило нет", "как правило нет", "редко, по договорённости", "обычно постоплата 30–90 дн.", "Есть переторжка — закладываем торг"],
]: lg.append(r)
for i, w in enumerate([24, 30, 34, 26, 30, 50], 1): lg.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
tidy(wb, row_h=48)
ws2 = wb.create_sheet("С площадок")
ws2.append(["Номер", "Предмет", "Аванс", "Оплата", "Срок/место поставки", "Обеспечение заявки", "Обеспечение договора", "Гарантия", "Прочее", "Источник"])
for c in ws2[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")
for x, r in rows:
    w = WEB.get(x["num"])
    if not (w and w.get("found")): continue
    ws2.append([x["num"], short(x["name"], 90)] + [w.get(k) or "—" for k in ("advance", "payment", "delivery", "sec_app", "sec_contr", "warranty", "other", "source")])
    ws2.cell(ws2.max_row, 1).number_format = "@"
    ws2.row_dimensions[ws2.max_row].height = 200
for i, wd in enumerate([20, 30, 34, 30, 38, 24, 30, 22, 34, 40], 1): ws2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = wd
for row in ws2.iter_rows():
    for c in row: c.alignment = Alignment(wrap_text=True, vertical="top"); c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
ws2.row_dimensions[1].height = 30; ws2.freeze_panes = "C2"
ws2.page_setup.orientation = "landscape"; ws2.page_setup.fitToWidth = 1; ws2.page_setup.fitToHeight = 0; ws2.sheet_properties.pageSetUpPr.fitToPage = True
wb.save("../CONDITIONS_2510_1910.xlsx")
json.dump([{**{k: x[k] for k in ("num", "name", "law", "method", "nmck", "cust")}, **{k: v for k, v in r.items()}} for x, r in rows], open("data/cond_2510.json", "w"), ensure_ascii=False, indent=1, default=str)
print(len(rows), "лотов")
for x, r in rows: print(x["num"], "|", r["app"], "|", r["contr"], "|", r["adv"][:30], "|", r["term"], "|", r["pref"], "|", r["freeze"])
