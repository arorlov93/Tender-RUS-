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
