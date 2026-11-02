"""Новая партия закупок 02.11.2026 (9 выгрузок, data/zakupki_0211/*.xlsx): отсев и решение.
Правила ГИДРОПАС: аукционы не берём; услуги/работы не берём; профиль — трубопроводная арматура и смежное (трубы, фасонка, КИП).
Карточки топ-закупок 223-ФЗ — gosplan (v2test), сохранены в data/zakupki_0211/cards/*.json (разъяснения, редакции, сроки).
→ ../ZAKUPKI_0211.xlsx"""
import openpyxl, glob, re, json, os, collections
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from tidy import tidy

H = None; rows = {}
for f in sorted(glob.glob("data/zakupki_0211/z_*.xlsx")):
    ws = openpyxl.load_workbook(f, data_only=True).active
    for i, r in enumerate(ws.iter_rows(values_only=True)):
        if i == 1: H = list(r); H[25] = "Заказчик"
        if i >= 2 and r[0]: rows.setdefault(str(r[0]), dict(zip(H, r)))
R = list(rows.values())
AUC = re.compile(r"аукцион", re.I)
SERV = re.compile(r"оказани|услуг|работ[ыа]?\b|выполнени|ремонт|монтаж|обслуживан|строительств|разработк|проектир|аренд|перевозк|охран|уборк|клининг|страхован|обучени|организаци|экспертиз|консульт|аудит|оценк|содержани", re.I)
CAT = [("Арматура трубопроводная", r"трубопроводн.{0,15}арматур|запорн.{0,15}арматур|арматур[аы]? (высокого|среднего|низкого)|задвиж|клапан|кран[аы]? шаров|затвор|вентил|ЗРА|компенсатор|манометр|арматур.{0,20}(АЗ|энергоблок|судов)"),
       ("Линейная/кабельная арматура (СИП, ЛЭП)", r"арматур.{0,30}(СИП|самонесущ|ЛЭП|линейн|кабельн|контактн)|линейн.{0,10}арматур|кабельн.{0,10}арматур"),
       ("Трубы, фасонка, ППУ", r"труб|фасонн|ППУ|фитинг|отвод|тройник|фланц"),
       ("Кабель, провод", r"кабел|провод|СИП"),
       ("Электротехника", r"электротех|электрооборуд|автомат|светильн|трансформатор|муфт|изолятор|счетчик|ввод|щит"),
       ("Насосы, приводы, промоборудование", r"насос|электродвигател|редуктор|подшипник|привод|оборудован|ЗИП|запасн|запчаст"),
       ("Металл, метизы", r"металл|прокат|метиз|сталь"),
       ("Медицина", r"медицин|лекарств|препарат|реагент|реактив|шприц|пробир|стоматолог|эндопротез|сустав"),
       ("Шины, авто, ГСМ", r"шин[аы]|автомобил|транспортн.{0,3}средств|спецтехник|масл|топлив|ГСМ"),
       ("IT, связь, оргтехника", r"компьютер|ноутбук|сервер|програм|лицензи|оргтехник|картридж|связ|сетев|ИБП|коммутатор")]
def cat(t):
    for c, rx in CAT:
        if re.search(rx, t, re.I): return c
    return "Прочее"
for r in R:
    t = str(r["Название"] or ""); r["cat"] = cat(t)
    r["auc"] = bool(AUC.search(str(r["Способ отбора"]))); r["serv"] = bool(SERV.search(t)) and r["cat"] not in ("Арматура трубопроводная",)
    try: r["nmc"] = float(r["НМЦ"]) if r["НМЦ"] else None
    except (TypeError, ValueError): r["nmc"] = None

# ---- решения по профильным (разобраны вручную по карточкам ЕИС/разъяснениям)
D = json.load(open("data/zakupki_0211/decisions.json"))
thin = Side(style="thin", color="BFBFBF"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
F = {"ИДЁМ": "C6EFCE", "СРОЧНО СМОТРЕТЬ ТЗ": "FFEB9C", "СМОТРЕТЬ ТЗ": "FFF2CC", "ЗАПРОСИТЬ / B2B": "DDEBF7", "НЕТ": "F8CBAD"}
def hd(ws):
    for c in ws[ws.max_row]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True, vertical="center")
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Решение по профильным"
ws.append(["Новая партия закупок 02.11.2026 — решение по профильным (арматура, трубы). Аукционы и услуги отсеяны по правилам компании."]); ws["A1"].font = Font(bold=True, size=12)
ws.append(["Решение", "№ закупки", "Заказчик", "Предмет", "НМЦ, ₽", "Окончание подачи", "Способ", "Площадка", "Почему", "Что сделать"]); hd(ws)
for d in D:
    r = rows.get(d["n"], {})
    ws.append([d["dec"], d["n"], d.get("cust") or r.get("Заказчик"), d.get("subj") or r.get("Название"), r.get("nmc") or d.get("nmc"), str(d.get("close") or r.get("Окончание приема заявок") or "")[:16],
               r.get("Способ отбора"), r.get("ЭТП"), d["why"], d["todo"]])
    X = ws.max_row; ws.cell(X, 1).fill = PatternFill("solid", fgColor=F[d["dec"]]); ws.cell(X, 5).number_format = "#,##0"
    for c in ws[X]: c.border = BOX; c.alignment = Alignment(wrap_text=True, vertical="top")
for i, w in enumerate([16, 18, 26, 46, 15, 15, 18, 16, 60, 50], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "C3"

# ---- отбор до 100 млн, ожидаемая маржа > 20 % (оценка по категории товара, способу и заказчику; ТЗ не скачивались — ЕИС недоступен)
MG = json.load(open("data/zakupki_0211/margin20.json"))
w0 = wb.create_sheet("Маржа >20 % до 100 млн", 0)
w0.append(["Отбор до 100 млн ₽ (кроме аукционов и услуг), где ожидаемая маржа выше 20 %. Маржа — оценка: (НМЦ − закупка с НДС) / НМЦ по опыту категории, без доставки и до торга"]); w0["A1"].font = Font(bold=True, size=12)
w0.append(["Группа", "№ закупки", "Заказчик", "Предмет", "НМЦ, ₽", "Подача до", "Способ", "Площадка", "Ожидаемая маржа", "Почему маржа", "Риск / что проверить"]); hd(w0)
grp = None
for c, n, m, why, risk in MG:
    r = rows[n]
    w0.append([c, n, r["Заказчик"], r["Название"], r["nmc"], str(r["Окончание приема заявок"] or "")[:16], r["Способ отбора"], r["ЭТП"], m, why, risk])
    X = w0.max_row; w0.cell(X, 5).number_format = "#,##0"; w0.cell(X, 9).fill = PatternFill("solid", fgColor="C6EFCE" if m.startswith(("25", "30")) else "FFF2CC")
    for cc in w0[X]: cc.border = BOX; cc.alignment = Alignment(wrap_text=True, vertical="top")
w0.append(["ИТОГО", len(MG), "", "", sum(rows[n]["nmc"] or 0 for _, n, *_ in MG)]); w0.cell(w0.max_row, 5).number_format = "#,##0"; [setattr(cc, "font", Font(bold=True)) for cc in w0[w0.max_row]]
w0.append([]); w0.append(["Не взяли (маржа обычно ниже 20 %): кабель и провод (5–12 %), трубы и ППУ-фасонка (3–7 %), серверы/IT (8–15 %), техническая соль, лекарства (рег. цены), высоковольтные вводы (только напрямую с завода), продажи имущества (Хевел). ДГК 186 млн и ТГК-1 1,2 млрд — выше лимита 100 млн."])
for i, wd in enumerate([22, 20, 28, 50, 14, 15, 16, 16, 11, 55, 40], 1): w0.column_dimensions[openpyxl.utils.get_column_letter(i)].width = wd
w0.freeze_panes = "C3"; wb.active = 0

# ---- доскональная проверка: позиции и цены 44-ФЗ (извещения), карточки и прошлые договоры заказчиков 223-ФЗ, рынок
V = json.load(open("data/zakupki_0211/verdict.json"))
wv = wb.create_sheet("Проверка — выгодно или нет", 0)
wv.append(["Проверка каждой закупки с потенциалом (до 100 млн): позиции и цены из извещений ЕИС (44-ФЗ), карточки и договоры заказчиков (223-ФЗ), рыночные цены. «РАСЧЁТ» — посчитано по позициям; «ОЦЕНКА» — ТЗ не скачан (zakupki.gov.ru недоступен), нужен файл ТЗ"]); wv["A1"].font = Font(bold=True, size=12)
wv.append(["Вердикт", "№ закупки", "Заказчик", "Предмет", "НМЦ, ₽", "Подача до", "Площадка", "Маржа", "Расчёт / основание", "Что делать"]); hd(wv)
VF = {"ВЫГОДНО — РАСЧЁТ": "C6EFCE", "ВЫГОДНО — ОЦЕНКА, НУЖЕН ТЗ": "E2EFDA", "НЕВЫГОДНО — РАСЧЁТ": "F8CBAD", "НЕВЫГОДНО — ОЦЕНКА": "FCE4D6", "НЕ УСПЕВАЕМ": "D9D9D9"}
order = list(VF)
for n, v, m, why, todo in sorted(V, key=lambda x: (order.index(x[1]), -(rows[x[0]]["nmc"] or 0))):
    r = rows[n]
    wv.append([v, n, r["Заказчик"], r["Название"], r["nmc"], str(r["Окончание приема заявок"] or "")[:16], r["ЭТП"], m, why, todo])
    X = wv.max_row; wv.cell(X, 1).fill = PatternFill("solid", fgColor=VF[v]); wv.cell(X, 5).number_format = "#,##0"
    for cc in wv[X]: cc.border = BOX; cc.alignment = Alignment(wrap_text=True, vertical="top")
wv.append([])
for v in order:
    L = [x for x in V if x[1] == v]; wv.append([v, len(L), "", "", sum(rows[x[0]]["nmc"] or 0 for x in L)]); wv.cell(wv.max_row, 5).number_format = "#,##0"; wv.cell(wv.max_row, 1).font = Font(bold=True)
for i, wd in enumerate([24, 20, 26, 46, 14, 15, 16, 10, 70, 42], 1): wv.column_dimensions[openpyxl.utils.get_column_letter(i)].width = wd
wv.freeze_panes = "C3"; wb.active = 0

s2 = wb.create_sheet("Отсев — сводка")
s2.append(["Категория", "Всего", "Аукционы (не идём)", "Услуги/работы (не идём)", "Товар, не аукцион", "НМЦ товара не аукцион, млн ₽", "из них с указанной НМЦ"]); hd(s2)
agg = collections.defaultdict(lambda: [0, 0, 0, 0, 0.0, 0])
for r in R:
    a = agg[r["cat"]]; a[0] += 1
    if r["auc"]: a[1] += 1
    elif r["serv"]: a[2] += 1
    else: a[3] += 1; a[4] += (r["nmc"] or 0) / 1e6; a[5] += r["nmc"] is not None
for k, a in sorted(agg.items(), key=lambda x: -x[1][0]): s2.append([k, *a[:4], round(a[4], 1), a[5]])
s2.append(["ИТОГО", *[sum(a[i] for a in agg.values()) for i in range(4)], round(sum(a[4] for a in agg.values()), 1), sum(a[5] for a in agg.values())])
for c in s2[s2.max_row]: c.font = Font(bold=True)
for i, w in enumerate([40, 8, 14, 16, 14, 18, 14], 1): s2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

def listing(title, flt):
    w = wb.create_sheet(title)
    w.append(["№ закупки", "Категория", "Предмет", "НМЦ, ₽", "Окончание подачи", "Способ", "Тип", "Площадка", "Заказчик", "Регион", "Место поставки", "Решение"]); hd(w)
    dec = {d["n"]: d["dec"] for d in D}
    for r in sorted([r for r in R if flt(r)], key=lambda r: (-(r["nmc"] or 0), str(r["Окончание приема заявок"]))):
        w.append([r["Номер"], r["cat"], r["Название"], r["nmc"], str(r["Окончание приема заявок"] or "")[:16], r["Способ отбора"], r["Тип торгов"], r["ЭТП"], r["Заказчик"], r["Регион"],
                  str(r["Место поставки"] or "")[:120], dec.get(str(r["Номер"]), "АУКЦИОН — не идём" if r["auc"] else "услуги — не идём" if r["serv"] else "")])
        w.cell(w.max_row, 4).number_format = "#,##0"
    for i, wd in enumerate([20, 22, 60, 14, 15, 20, 12, 16, 30, 18, 30, 18], 1): w.column_dimensions[openpyxl.utils.get_column_letter(i)].width = wd
    w.freeze_panes = "C2"; w.auto_filter.ref = f"A1:L{w.max_row}"
listing("Профиль — арматура, трубы", lambda r: r["cat"] in ("Арматура трубопроводная", "Трубы, фасонка, ППУ") and not r["serv"])
listing("Кабель, электро, оборуд.", lambda r: r["cat"] in ("Кабель, провод", "Электротехника", "Насосы, приводы, промоборудование", "Линейная/кабельная арматура (СИП, ЛЭП)") and not r["auc"] and not r["serv"])
listing("Все закупки", lambda r: True)
tidy(wb, row_h=32)
wb.save("../ZAKUPKI_0211.xlsx")
print({k: v[:4] for k, v in agg.items()})
