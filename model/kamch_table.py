"""Камчатскэнерго 32616346943 — арматура: позиции, количества, себестоимость (если есть price_*.json), маржа к НМЦК, условия, история."""
import json, os, re, sys, openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from tidy import tidy

PR = sys.argv[1] if len(sys.argv) > 1 else ""
NMCK = 101102908
Q = {(x["table"], x["n"]): x for x in json.load(open("data/kamch/kam_qty.json"))}
S = json.load(open("data/kamch/kam_spec.json"))["sheet2"]
P = {}
for g in "ABC":
    f = f"{PR}/price_{g}.json"
    if PR and os.path.exists(f):
        for x in json.load(open(f)):
            m = re.search(r"1\.\d", str(x.get("spec", "")))
            if m: P[(m.group(0), int(x["n"]))] = x
K = json.load(open("data/big/big_kamch.json"))[0]
GRP = {}
for g, nm in (("A", "Задвижки и затворы"), ("B", "Клапаны"), ("C", "Краны шаровые и прочее")):
    f = f"{PR}/grp_{g}.json"
    if PR and os.path.exists(f):
        for x in json.load(open(f)):
            GRP[(re.search(r"1\.\d", x["spec"]).group(0), x["n"])] = nm
TABLES = {"1.1": "Спец. №1 · Возобновляемая энергетика · ремонт", "1.2": "Спец. №2 · КТЭЦ · ремонт арендованного", "1.3": "Спец. №2 · КТЭЦ · ремонт",
          "1.4": "Спец. №2 · КТЭЦ · эксплуатация", "1.5": "Спец. №2 · КТЭЦ · ТО", "1.6": "Спец. №3 · Коммунальная энергетика · ремонт",
          "1.7": "Спец. №3 · Коммунальная энергетика · ТО", "1.8": "Спец. №4 · ЦЭС · ремонт"}
FREIGHT = 55  # ₽/кг до Петропавловска-Камчатского (оценка: 45–60 ₽/кг, контейнер морем через Владивосток)
TORG = 0.10   # торг к ценам сайтов/прайсов
thin = Side(style="thin", color="BFBFBF"); M = '#,##0" ₽"'; PCT = "0.0%"


def head(ws, r):
    for c in ws[r]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")


wb = openpyxl.Workbook()
sv = wb.active; sv.title = "Итог"
ws = wb.create_sheet("Позиции")
ws.append(["Камчатскэнерго 32616346943 — позиции, количества, себестоимость (заполненные цены — рынок 2026 с НДС; пустые — ещё ищем, можно вписать вручную)"]); ws["A1"].font = Font(bold=True, size=12)
H = ["Таблица", "№", "Группа", "Позиция", "Кол-во", "Ед.", "DN", "PN, МПа", "Материал корпуса", "Привод", "Цена за ед. (рынок), ₽", "Цена после торга −10%, ₽", "Вес ед., кг",
     "Сумма закупки, ₽", "Доставка, ₽", "Себестоимость позиции, ₽", "Себес за ед. (торг+доставка), ₽", "Поставщик", "Тип цены / подтверждение 2026", "Аналог / реестр МПТ", "Ссылка"]
ws.append(H); head(ws, 2)
first = 3
for x in S:
    m = re.search(r"Таблица\s*(1\.\d)", x["spec"]); t = m.group(1); n = x["n"]
    q = Q[(t, n)]; p = P.get((t, n), {})
    r = ws.max_row + 1
    ws.append([t, n, GRP.get((t, n), ""), q["name"], q["qty"], q["unit"], x.get("dn"), x.get("pn"), str(x.get("mat") or "")[:40], str(x.get("drive") or "")[:30],
               p.get("price"), f"=IF(K{r}=\"\",\"\",K{r}*(1-{TORG}))", p.get("weight_kg"),
               f"=IF(K{r}=\"\",\"\",L{r}*E{r})", f"=IF(M{r}=\"\",\"\",M{r}*E{r}*{FREIGHT})", f"=IF(N{r}=\"\",\"\",N{r}+N(O{r}))", f"=IF(P{r}=\"\",\"\",P{r}/E{r})",
               str(p.get("supplier") or "")[:35], " · ".join(str(v) for v in (p.get("price_type"), p.get("date_evidence")) if v)[:55],
               " · ".join(str(v) for v in (p.get("analog"), p.get("registry")) if v)[:55], p.get("url")])
    if p.get("url"): ws.cell(r, 21).hyperlink = p["url"]
    for j in (11, 12, 14, 15, 16, 17): ws.cell(r, j).number_format = M
    if not p.get("price"): ws.cell(r, 11).fill = PatternFill("solid", fgColor="FFF2CC")
last = ws.max_row
ws.append(["", "", "", "ИТОГО", f"=SUM(E{first}:E{last})", "", "", "", "", "", "", "", "", f"=SUM(N{first}:N{last})", f"=SUM(O{first}:O{last})", f"=SUM(P{first}:P{last})"])
for j in (14, 15, 16): ws.cell(ws.max_row, j).number_format = M
for c in ws[ws.max_row]: c.font = Font(bold=True)
for row in ws.iter_rows(min_row=3):
    for c in row: c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
for i, w in enumerate([7, 5, 16, 46, 7, 5, 6, 7, 18, 16, 13, 13, 8, 14, 12, 14, 13, 22, 26, 26, 30], 1):
    ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "E3"; ws.auto_filter.ref = f"A2:U{last}"
TOT = f"Позиции!P{last + 1}"; CNT = f"COUNT(Позиции!K{first}:K{last})"

# Итог
sv.append(["Камчатскэнерго 32616346943 — арматура среднего и низкого давления (Лот 0200-РЕМ ПРОД-2027-КамчЭн)"]); sv["A1"].font = Font(bold=True, size=13)
sv.append(["Конкурс в эл. форме, 223-ФЗ, Росэлторг (секция РусГидро). Подача до 29.10.2026. НМЦК с НДС. Наша цена = цена договора (УСН)."])
# --- себестоимость по группам (значения считаются здесь же; лист «Позиции» — формулами для ручной правки)
G = {}
for x in S:
    t = re.search(r"Таблица\s*(1\.\d)", x["spec"]).group(1); k = (t, x["n"]); g = GRP.get(k, "Прочее"); q = Q[k]["qty"] or 0; p = P.get(k, {})
    d = G.setdefault(g, dict(pos=0, priced=0, qty=0, mkt=0.0, frt=0.0, w=0.0))
    d["pos"] += 1; d["qty"] += q
    if p.get("price"):
        d["priced"] += 1; d["mkt"] += p["price"] * q; d["frt"] += (p.get("weight_kg") or 0) * q * FREIGHT; d["w"] += (p.get("weight_kg") or 0) * q
HIST = K["cost_total"] + K["logistics"]
for d in G.values():
    d["torg"] = -d["mkt"] * TORG; d["cost"] = d["mkt"] + d["torg"] + d["frt"]
    d["est"] = d["cost"] / d["priced"] * (d["pos"] - d["priced"]) if 0 < d["priced"] < d["pos"] else 0.0
miss = [g for g, d in G.items() if d["priced"] == 0]
known = sum(d["cost"] + d["est"] for d in G.values())
for g in miss: G[g]["est"] = max(HIST - known, 0) / len(miss)
TOTAL = sum(d["cost"] + d["est"] for d in G.values())
sv.append(["Себестоимость — из чего состоит", "", "", "", "", "", "", "", ""]); head(sv, 3)
sv.append(["Группа", "Позиций", "С ценой", "Закупка по рынку (сайты/прайсы 2026, с НДС), ₽", "Торг −10%, ₽", "Доставка до Камчатки (55 ₽/кг), ₽",
           "Себестоимость найденных, ₽", "Оценка ненайденных, ₽", "Себестоимость группы, ₽"]); head(sv, 4)
for g in ("Задвижки и затворы", "Клапаны", "Краны шаровые и прочее", "Прочее"):
    if g not in G: continue
    d = G[g]
    sv.append([g, d["pos"], d["priced"], d["mkt"], d["torg"], d["frt"], d["cost"], d["est"], d["cost"] + d["est"]])
sv.append(["ИТОГО", sum(d["pos"] for d in G.values()), sum(d["priced"] for d in G.values()), sum(d["mkt"] for d in G.values()), sum(d["torg"] for d in G.values()),
           sum(d["frt"] for d in G.values()), sum(d["cost"] for d in G.values()), sum(d["est"] for d in G.values()), TOTAL])
for c in sv[sv.max_row]: c.font = Font(bold=True)
for r in range(5, sv.max_row + 1):
    for j in range(4, 10): sv.cell(r, j).number_format = M
sv.append(["Как считали", "Себестоимость = цена поставщика за ед. × кол-во − 10% торг + вес × кол-во × 55 ₽/кг. Ненайденные позиции: по средней найденной в группе; если в группе цен нет — "
           "остаток до оценки по истории лотов (60% НМЦК + 5 млн доставки). Детально по каждой позиции — лист «Позиции» (столбец «Себес за ед.»)."])
sv.append([])
sv.append(["Показатель", "Значение", "Комментарий"]); head(sv, sv.max_row)
r0 = sv.max_row + 1
rows = [
    ("НМЦК, ₽", NMCK, "весь лот, 315 позиций, 3 383 ед., 4 грузополучателя"),
    ("Себестоимость лота, ₽", TOTAL, "из таблицы выше"),
    ("Маржа при НМЦК", f"=1-B{r0 + 1}/B{r0}", ""),
    ("Ожидаемое снижение на конкурсе", 0.15, "история: −7% … −23% (2025–2026), большинство заявок отклоняют"),
    ("Цена заявки, ₽", f"=B{r0}*(1-B{r0 + 3})", ""),
    ("Прибыль до налога, ₽", f"=B{r0 + 4}-B{r0 + 1}-0.008*B{r0 + 4}-B{r0 + 1}*0.2*120/365", "0,8% ЭП/банк, деньги 20% годовых на 120 дней"),
    ("Прибыль после УСН 15%, ₽", f"=B{r0 + 5}-0.15*MAX(B{r0 + 5},0)", ""),
    ("Маржа к цене заявки", f"=B{r0 + 6}/B{r0 + 4}", ""),
]
for r in rows: sv.append(list(r))
for k in (0, 1, 4, 5, 6): sv.cell(r0 + k, 2).number_format = M
for k in (2, 3, 7): sv.cell(r0 + k, 2).number_format = PCT
B4, B7 = f"B{r0}", f"B{r0 + 1}"
sv.append([])
sv.append(["Сценарии снижения", "Прибыль после УСН, ₽", ""]); head(sv, sv.max_row)
for d in (0, 0.07, 0.15, 0.23, 0.30):
    r = sv.max_row + 1
    e = f"({B4}*(1-{d})-{B7}-0.008*{B4}*(1-{d})-{B7}*0.2*120/365)"
    sv.append([f"−{int(d * 100)}%", f"={e}*IF({e}>0,0.85,1)", ""])
    sv.cell(r, 2).number_format = M
sv.append([])
sv.append(["Условия", "", ""]); head(sv, sv.max_row)
for a, b in [("Поставка", "в течение 120 календарных дней с подписания договора; Петропавловск-Камчатский (ул. Пограничная 75А, ул. Степная 50, пр. Победы 24) и Елизовский р-н, пос. Светлый"),
             ("Обеспечение заявки", "1 005 499,42 ₽ (БГ или деньги)"), ("Аванс", "предусмотрен; обеспечение возврата аванса — 100% суммы аванса (БГ); размер — в проекте договора"),
             ("Гарантия", "по заводу, не менее 12 мес.; изготовление — год поставки или предыдущий"),
             ("Нацрежим", "ПП 1875 — ограничение + преимущество: по каждой позиции номер записи в реестре Минпромторга/ЕАЭС"),
             ("Документы", "сертификаты/декларации ТР ТС 032/2013 и 010/2011 со схемами 1с/5Д и протоколами испытаний; техпредложение по Прил. 1"),
             ("Аккредитация", "аккредитация в РусГидро (Прил. 10)"), ("Подача на часть лота", "нельзя — лот неделимый"),
             ("Эквиваленты", "допускаются, марка носит описательный характер; нужно обоснование эквивалентности")]:
    sv.append([a, b, ""])
sv.append([])
sv.append(["История лотов Камчатскэнерго", "", ""]); head(sv, sv.max_row)
for h in K.get("history") or []:
    if isinstance(h, dict):
        sv.append([h.get("lot") or h.get("num"), f'НМЦК {h.get("nmck"):,.0f} ₽ → победитель {h.get("winner_price") or "—"}; снижение {h.get("disc")}'.replace(",", " "),
                   str(h.get("rejected") or h.get("contract") or "")[:150]])
    else: sv.append([str(h)[:60], "", ""])
for row in sv.iter_rows(min_row=3):
    for c in row: c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
for i, w in enumerate([34, 16, 12, 22, 16, 18, 18, 18, 20], 1): sv.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

# По таблицам
st = wb.create_sheet("По таблицам")
st.append(["Таблица", "Грузополучатель / статья", "Позиций", "Единиц", "Цен найдено", "Себестоимость найденных, ₽"]); head(st, 1)
for t, nm in TABLES.items():
    r = st.max_row + 1
    st.append([t, nm, f'=COUNTIF(Позиции!A{first}:A{last},"{t}")', f'=SUMIF(Позиции!A{first}:A{last},"{t}",Позиции!E{first}:E{last})',
               f'=COUNTIFS(Позиции!A{first}:A{last},"{t}",Позиции!K{first}:K{last},">0")', f'=SUMIF(Позиции!A{first}:A{last},"{t}",Позиции!P{first}:P{last})'])
    st.cell(r, 6).number_format = M
for row in st.iter_rows():
    for c in row: c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
for i, w in enumerate([9, 48, 10, 10, 12, 22], 1): st.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

tidy(wb, row_h=34)
for r in range(r0 - 1, r0 + 8):
    sv.merge_cells(start_row=r, start_column=3, end_row=r, end_column=9)
    sv.cell(r, 3).alignment = Alignment(wrap_text=True, vertical="center")
for r in range(1, sv.max_row + 1):
    if r0 - 1 <= r < r0 + 8: continue
    v = sv.cell(r, 2).value
    if isinstance(v, str) and not v.startswith("=") and len(v) > 25 and sv.cell(r, 4).value in (None, ""):
        sv.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9)
        sv.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="center")
        sv.row_dimensions[r].height = 15 * (len(v) // 120 + 1) + 8
wb.save("../KAMCHATKA_32616346943.xlsx")
print("позиций", len(S), "с ценой", len(P), "себес", round(TOTAL), {g: (d["priced"], round(d["cost"] + d["est"])) for g, d in G.items()})
