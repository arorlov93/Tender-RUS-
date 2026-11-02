"""Лучшие закупки из всех 9 выгрузок партии 02.11 (подача с 06.11, не аукционы, товары; включая запросы КП/цен/мониторинг).
570 закупок разобраны по одной (deep/all570_*_out.json: группа, оценка маржи по нашему опыту, состав по карточкам gosplan/b2b, риски, шаг).
Где есть ТЗ и расчёт по ценам — оценка заменена расчётом (глубокая проверка, deep/*.json). → ../LUCHSHIE_0211.xlsx"""
import json, re, os, openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from tidy import tidy

D = "data/zakupki_0211/tz/deep"
RAW = {r[0]: r for r in json.load(open("data/zakupki_0211/keep0611c.json"))}
OUT = [x for i in range(1, 5) for x in json.load(open(f"{D}/all570_{i}_out.json"))]
# Расчёт по ТЗ (перекрывает оценку): номер → (маржа, при −10 %, срок поставки, вывод, комментарий)
CALC = {
    "32616526616": (0.60, 0.556, "по заявкам 15 раб. дн.; период 01.01–31.12.2027", "МАРЖА ЕСТЬ, СРОК > 6 мес", "расчёт по 47 поз. (84 % НМЦ)"),
    "32616475914": (0.347, 0.274, "по заявкам ≤130 дн.; до 31.12.2027", "МАРЖА ЕСТЬ, СРОК > 6 мес", "расчёт по 56 поз. (85 % НМЦ); лимит 4,27 млн"),
    "32616488988": (0.261, 0.179, "30 дн. с договора", "ТОНКО", "глубокая проверка 208/222 поз."),
    "32616531898": (0.121, 0.023, "30 дн. с договора", "НЕ ИДЁМ (без скидки Промрукава ≥15 %)", "ПП 719: гофротрубы только Нептун/Промрукав"),
    "32616539033": (-0.168, -0.298, "по заявкам 5 раб. дн., до 31.12.2026", "НЕ ИДЁМ (нужен опт −25–30 %)", "розница фабрик"),
    "32616485509": (0.088, -0.013, None, "НЕ ИДЁМ", "мин. цена из обоснования НМЦ"),
    "32616459808": (-0.068, -0.187, None, "НЕ ИДЁМ", ""),
    "1038631": (0.203, 0.114, None, "НЕ ИДЁМ", "расчёт по 4 поз."),
    "B2910261728004": (None, None, None, "НЕ ИДЁМ", "21/22 поз. = 85 % НМЦ"),
    "32616524948": (None, None, None, "НЕ ИДЁМ", "9/10 поз. = 75 % НМЦ"),
    "B2810261718575": (None, None, None, "НЕ ИДЁМ", "34/46 поз. = 62 % НМЦ"),
    "32616465624": (None, None, None, "НЕ ИДЁМ", "12/17 поз. = 49 % НМЦ без FKM/PTFE и фланцев"),
    "32616535058": (0.08, -0.02, None, "НЕ ИДЁМ", ""),
    "32616530999": (None, None, None, "НУЖНЫ КП", "12 импортных поз. без рублёвых цен"),
    "5590355": (None, None, None, "НУЖНЫ КП", "ОСТ5 с ОТК — только заводы"),
    "3511105": (None, None, None, "НУЖНЫ КП", ""),
    "ГПК611061": (None, None, None, "НУЖНЫ КП", ""),
}
if os.path.exists(f"{D}/calc_new.json"):
    for k, v in json.load(open(f"{D}/calc_new.json")).items(): CALC[k] = tuple(v)
GAZ = lambda n: n in {"32616420912", "1038513", "1038631"} or (n != "ГПК611061" and bool(re.search(r"Газпром|ТГК-1|Петербургтеплоэнерго|Мосэнерго|ГЭХ", str((RAW.get(n) or [""] * 26)[25]) + str((RAW.get(n) or [""] * 3)[2]))))
def cls(x):
    n = x["n"]
    if n in CALC:
        m10, v = CALC[n][1], CALC[n][3]
        if m10 is not None and m10 >= 0.25 and ("СРОК" not in v or GAZ(n)): return 0  # Газпром — берём и при договоре на 2027 г.
        if v.startswith(("МАРЖА ЕСТЬ", "ТОНКО", "НУЖНЫ")): return 2
        return 4
    v = x["verdict"]
    return 1 if v.startswith("ЛУЧШАЯ") else 2 if v.startswith("ВЫГОДНО") else 3 if v.startswith("ПОД") else 4
lo = lambda s: (lambda m: int(m.group(1)) if m else 0)(re.search(r"(\d+)\s*[–-]", str(s)))
nmc = lambda n: float(RAW[n][3] or 0)
OUT.sort(key=lambda x: (cls(x), -lo(x["margin_est"]), -nmc(x["n"])))
CL = {0: ("ИДЁМ — расчёт ≥ 25 %", "63BE7B"), 1: ("ЛУЧШАЯ — оценка", "C6EFCE"), 2: ("ВЫГОДНО / нужен ТЗ или КП", "FFF2CC"), 3: ("ПОД ВОПРОСОМ", "FCE4D6"), 4: ("НЕТ", "F8CBAD")}

thin = Side(style="thin", color="BFBFBF"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Рейтинг 570"
ws.append(["Все закупки 9 выгрузок с подачей с 06.11 (не аукционы, товары; вкл. запросы КП/цен) — рейтинг. Маржа: «расчёт» — по ТЗ и рыночным ценам; «оценка» — по нашему опыту (арматура у заводов 37–49 %, электротехника 26–60 %, КИП 35 %, соль 39–55 %, кабель 5–12 %)."]); ws["A1"].font = Font(bold=True, size=11)
H = ["Ранг", "Класс", "№ закупки", "Подача до", "НМЦ, ₽", "Способ", "Площадка", "Запрос КП", "Заказчик", "Регион", "Предмет", "Группа", "Маржа при −10 %", "Чем подтверждена", "Срок поставки", "Что поставляется", "Риски", "Вывод", "Что делать", "Ссылка"]
ws.append(H)
for c in ws[2]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True, vertical="top")
top = []
for i, x in enumerate(OUT, 1):
    r = RAW[x["n"]]; k = cls(x); c = CALC.get(x["n"])
    mg = (f"{c[1]:.0%}" if c[1] is not None else "—") if c else x["margin_est"]
    src = ("расчёт по ТЗ: " + (c[4] or "")) if c else "оценка: " + str(x.get("basis") or "")[:200]
    term = (c[2] if c and c[2] else x.get("term"))
    verdict = c[3] if c else x["verdict"]
    ws.append([i, CL[k][0], x["n"], str(r[13])[:16], nmc(x["n"]) or "не указана", r[18], r[19], "да" if x.get("kp_request") else "", r[25], r[24], r[2], x.get("group"), mg, src, term,
               str(x.get("items_hint") or "")[:500], str(x.get("risks") or "")[:300], verdict, x.get("next_step"), x.get("url") or (f"https://zakupki.gov.ru/epz/order/notice/notice223/common-info.html?regNumber={x['n']}" if re.fullmatch(r"3\d{10}", x["n"]) else f"по номеру на площадке {r[19]}")])
    X = ws.max_row; ws.cell(X, 2).fill = PatternFill("solid", fgColor=CL[k][1]); ws.cell(X, 5).number_format = "#,##0"
    for cc in ws[X]: cc.border = BOX; cc.alignment = Alignment(wrap_text=True, vertical="top")
    if k <= 2: top.append((i, k, x, r, mg, term, verdict))
for j, w in enumerate([5, 16, 17, 14, 13, 14, 14, 7, 26, 14, 44, 18, 10, 30, 18, 40, 30, 24, 30, 30], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
ws.freeze_panes = "D3"; ws.auto_filter.ref = f"A2:T{ws.max_row}"
s0 = wb.create_sheet("ЛУЧШИЕ", 0)
s0.append(["Лучшие закупки из всех 9 выгрузок (подача с 06.11). Сначала — подтверждённые расчётом по ТЗ, затем лучшие по оценке (нужен ТЗ / КП заводов), затем «выгодно»."]); s0["A1"].font = Font(bold=True, size=12)
s0.append(["Ранг", "Класс", "№ закупки", "Подача до", "НМЦ, ₽", "Заказчик", "Предмет", "Маржа при −10 %", "Срок поставки", "Вывод", "Что делать"])
for c in s0[2]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True, vertical="top")
for i, k, x, r, mg, term, verdict in top:
    s0.append([i, CL[k][0], x["n"], str(r[13])[:16], nmc(x["n"]) or "не указана", r[25], r[2], mg, term, verdict, x.get("next_step")])
    X = s0.max_row; s0.cell(X, 2).fill = PatternFill("solid", fgColor=CL[k][1]); s0.cell(X, 5).number_format = "#,##0"
    for cc in s0[X]: cc.border = BOX; cc.alignment = Alignment(wrap_text=True, vertical="top")
for j, w in enumerate([5, 18, 17, 14, 13, 28, 50, 10, 22, 24, 40], 1): s0.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
s0.freeze_panes = "C3"

# ---------- Аукционы (по просьбе: включить аукционы тоже)
AR = {x["n"]: x for i in (1, 2) for x in json.load(open(f"{D}/auc_{i}.json"))}
AO = [x for i in (1, 2) for x in json.load(open(f"{D}/auc_{i}_out.json"))]
acl = lambda x: 1 if x["verdict"].startswith("ЛУЧШАЯ") else 2 if x["verdict"].startswith("ВЫГОДНО") else 3 if x["verdict"].startswith("ПОД") else 4
AO.sort(key=lambda x: (acl(x), -lo(x["margin_est"]), -float(AR[x["n"]]["nmc"] or 0)))
sa = wb.create_sheet("Аукционы", 2)
sa.append(["Аукционы с подачей с 06.11 (товары, 206 шт.). Маржа — при ОЖИДАЕМОМ снижении на торгах по группе (не −10 %); оценка, кроме позиций, сверенных с рынком по карточке ЕИС."]); sa["A1"].font = Font(bold=True, size=11)
sa.append(["Ранг", "Класс", "№ закупки", "Подача до", "НМЦ, ₽", "Способ", "Площадка", "Заказчик", "Регион", "Предмет", "Группа", "Ожидаемое снижение", "Маржа при снижении", "Чем подтверждена", "Срок поставки", "Что поставляется", "Риски", "Вывод", "Что делать", "Ссылка"])
for c in sa[2]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="7030A0"); c.alignment = Alignment(wrap_text=True, vertical="top")
for i, x in enumerate(AO, 1):
    r = AR[x["n"]]; k = acl(x)
    sa.append([i, CL[k][0].replace("ЛУЧШАЯ", "ЛУЧШАЯ (аукцион)"), x["n"], str(r["deadline"])[:16], float(r["nmc"] or 0) or "не указана", r["method"], r["etp"], r["customer"], r["region"], r["subject"], x.get("group"), x.get("expected_drop"), x.get("margin_est"),
               "оценка: " + str(x.get("basis") or "")[:200], x.get("term"), str(x.get("items_hint") or "")[:500], str(x.get("risks") or "")[:300], x["verdict"], x.get("next_step"), x.get("url")])
    X = sa.max_row; sa.cell(X, 2).fill = PatternFill("solid", fgColor=CL[k][1]); sa.cell(X, 5).number_format = "#,##0"
    for cc in sa[X]: cc.border = BOX; cc.alignment = Alignment(wrap_text=True, vertical="top")
    if k <= 2:
        s0.append([f"А{i}", "АУКЦИОН — " + CL[k][0], x["n"], str(r["deadline"])[:16], float(r["nmc"] or 0) or "не указана", r["customer"], r["subject"], f'{x.get("margin_est")} (снижение {x.get("expected_drop")})', x.get("term"), x["verdict"], x.get("next_step")])
        Y = s0.max_row; s0.cell(Y, 2).fill = PatternFill("solid", fgColor="D9C3E9"); s0.cell(Y, 5).number_format = "#,##0"
        for cc in s0[Y]: cc.border = BOX; cc.alignment = Alignment(wrap_text=True, vertical="top")
for j, w in enumerate([5, 18, 20, 14, 13, 16, 14, 26, 14, 44, 18, 10, 14, 30, 18, 40, 30, 24, 30, 30], 1): sa.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
sa.freeze_panes = "D3"; sa.auto_filter.ref = f"A2:T{sa.max_row}"

# ---------- ИТОГ: что подаём (вкл. группу Газпром — ГЭХ, ТГК-1, Газпром теплоэнерго, ТСК Мосэнерго)
si = wb.create_sheet("ИТОГ", 0)
si.append(["ИТОГ по 9 выгрузкам (подача с 06.11): что подаём. Включая группу Газпром (договоры на 2027 г. берём). Сортировка — по сроку подачи. «Расчёт» — по ТЗ и рыночным ценам; «оценка» — нужен ТЗ/КП заводов до подачи."]); si["A1"].font = Font(bold=True, size=12)
si.append(["Подача до", "№ закупки", "Тип", "Газпром", "Заказчик", "Предмет", "НМЦ, ₽", "Маржа", "Чем подтверждена", "Срок поставки", "Что делать до подачи", "Ссылка / площадка"])
for c in si[2]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True, vertical="top")
F = []
for x in OUT:
    n = x["n"]; r = RAW[n]; k = cls(x); c = CALC.get(n)
    if k <= 1 or (GAZ(n) and k <= 2):
        mg = (f"{c[1]:.0%} при −10 %" if c and c[1] is not None else x["margin_est"] + " при −10 %") if (c or x.get("margin_est")) else "—"
        F.append([str(r[13])[:16], n, "конкурс / запрос КП" if not x.get("kp_request") else "запрос КП / цен", "да" if GAZ(n) else "", r[25], r[2], nmc(n) or "не указана", mg,
                  "расчёт по ТЗ" if c else "оценка", (c[2] if c and c[2] else x.get("term")), (c[4] + "; " if c and c[4] else "") + str(x.get("next_step") or ""), x.get("url") or r[19], k])
for x in AO:
    if acl(x) == 1:
        r = AR[x["n"]]
        F.append([str(r["deadline"])[:16], x["n"], "АУКЦИОН", "да" if re.search("Газпром", str(r["customer"])) else "", r["customer"], r["subject"], float(r["nmc"] or 0) or "не указана", f'{x.get("margin_est")}', "оценка", x.get("term"), x.get("next_step"), x.get("url") or r["etp"], 1])
F.sort(key=lambda z: z[0])
for z in F:
    si.append(z[:12]); X = si.max_row
    si.cell(X, 7).number_format = "#,##0"
    col = "63BE7B" if z[12] == 0 else "C6EFCE" if z[12] == 1 else "FFF2CC"
    si.cell(X, 8).fill = PatternFill("solid", fgColor=col)
    if z[3]: si.cell(X, 4).fill = PatternFill("solid", fgColor="BDD7EE")
    if z[0] < "2026-11-07": si.cell(X, 1).fill = PatternFill("solid", fgColor="F8CBAD")
    for cc in si[X]: cc.border = BOX; cc.alignment = Alignment(wrap_text=True, vertical="top")
for j, w in enumerate([14, 18, 12, 8, 28, 48, 14, 20, 12, 22, 46, 30], 1): si.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
si.freeze_panes = "C3"; si.auto_filter.ref = f"A2:L{si.max_row}"
print("ИТОГ:", len(F))
tidy(wb, row_h=30); wb.save("../LUCHSHIE_0211.xlsx")
import collections; print(collections.Counter(cls(x) for x in OUT)); print(len(top))
