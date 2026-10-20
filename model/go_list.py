"""Куда заходим: выгодные лоты (без аукционов) со ссылками, перепроверенной себестоимостью и маржой. 20.10.2026."""
import json, os, sys, openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from tidy import tidy

U = json.load(open("data/units_1910.json"))
L = json.load(open("data/go_links_2010.json"))
V = json.load(open(sys.argv[1])) if len(sys.argv) > 1 and os.path.exists(sys.argv[1]) else {}
TC = {"цитрат": 9.9, "гепарин": 9.9, "активатор": 9.9, "K2ЭДТА": 9.9, "K3ЭДТА": 9.9}
TC.update({k: v for k, v in (V.get("tubes_recommended_cost") or {}).items() if isinstance(v, (int, float))})
TC["активатор"] = max(TC["активатор"], 15.0)  # гель не указан в извещении — берём дороже (с гелем 13,75 + доставка)
GORKA = V.get("gorka_recommended_cost") or 3500


def tube_cost(name, act=None):
    s = name.lower()
    if "цитрат" in s: return TC["цитрат"]
    if "гепарин" in s: return TC["гепарин"]
    if "k2" in s or "к2" in s: return TC["K2ЭДТА"]
    if "эдта" in s: return TC["K3ЭДТА"]
    return act or TC["активатор"]


def model(P, C, days=60):
    pre = P - C - 0.008 * P - C * 0.2 * days / 365
    return pre - 0.15 * max(pre, 0)


def be(nm, C, days=60):
    lo, hi = 0.0, 0.95
    for _ in range(60):
        m = (lo + hi) / 2
        lo, hi = (m, hi) if model(nm * (1 - m), C, days) > 0 else (lo, m)
    return lo


def tubes(num, extra, act=None):
    pos = U[num]; freight = extra * sum(p["qty"] for p in pos); C = sum(tube_cost(p["name"], act) * p["qty"] for p in pos) + freight
    return C, "; ".join(f'{p["qty"]:,.0f} шт × {tube_cost(p["name"], act):.1f} ₽ ({p["name"].split(" с ")[-1][:22]})'.replace(",", " ") for p in pos[:4]) + f"; доплата за дальнюю доставку {freight:,.0f} ₽ (МиниМед, прайс ©2026, РФ — импорт запрещён ПП 1875)".replace(",", " ")


def omsk(num):
    # зимний костюм 5 000 (правило клиента), утеплённые ботинки 2 618 (справочник 2026), сварщик зимний 5 990; доставка 35 ₽/кг
    m = {"Костюм для защиты от пониженных температур для рабочих": (5000, 3), "Костюм для защиты от пониженных температур для руководителей и специалистов": (5000, 3),
         "Ботинки утепленные": (2618, 1.8), "Костюм сварщика зимний": (5990, 4)}
    C = sum((m[p["name"]][0] + 35 * m[p["name"]][1]) * p["qty"] for p in U[num])
    return C, "; ".join(f'{p["name"][:28]} {p["qty"]:.0f}×{m[p["name"]][0]:,} ₽'.replace(",", " ") for p in U[num])


ROWS = []
def add(**k): ROWS.append(k)

C, b = tubes("0338200002226000430", 1.0)  # ~1,6 т до Камчатки авиа/море ~120 ₽/кг
add(num="0338200002226000430", cat="Пробирки", cust="ГБУЗ «Камчатский краевой онкологический диспансер»", subj="Пробирки вакуумные, 82 200 шт", end="23.10.2026 04:00",
    nmck=2221608, C=C, d=0.25, basis=b, law="44-ФЗ · запрос котировок", etp="Росэлторг",
    risk="активатор посчитан с гелем (15 ₽) — с запасом; нужны РУ и реестр РФ (импорт запрещён)")
C, b = tubes("0372200263426001564", 0.0, act=10.0)
add(num="0372200263426001564", cat="Пробирки", cust="СПб ГБУЗ «Городская Александровская больница»", subj="Пробирки вакуумные, 50 000 шт", end="23.10.2026 08:00",
    nmck=863700, C=C, d=0.25, basis=b, law="44-ФЗ · запрос котировок", etp="ЭТП ГПБ",
    risk="посчитано для активатора БЕЗ геля (10 ₽); если в ТЗ с гелем (15 ₽) — при −25% убыток ~29 тыс., подавать не ниже −20%")
C, b = tubes("0325300007726000124", 0.0)
add(num="0325300007726000124", cat="Пробирки", cust="ГБУЗ АО «Городская поликлиника №5», Астрахань", subj="Пробирки вакуумные K3ЭДТА, 30 000 шт", end="26.10.2026 08:00",
    nmck=496200, C=C, d=0.25, basis=b, law="44-ФЗ · запрос котировок", etp="Росэлторг")
add(num="32616484297", cat="Форма", cust="МАУ ДО ДМЦ «Алый парус», Тюменская обл.", subj="Костюм «Горка-3» летний, палатка, 121 шт", end="22.10.2026 05:00",
    nmck=671066, C=121 * GORKA, d=0.10, basis=f"121 шт × {GORKA:,.0f} ₽ с доставкой (Стоик «Горка СТ» опт 3 139 ₽, мин. заказ 50 тыс.)".replace(",", " "), law="223-ФЗ · электронный магазин", etp="РТС-тендер 223")
C, b = omsk("0352500002526000085")
add(num="0352500002526000085", cat="Спецодежда", cust="БУ Омской обл. «Тевризское ДРСУ»", subj="Спецодежда и обувь зимняя", end="22.10.2026 05:00",
    nmck=498744.79, C=C, d=0.10, basis=b, law="44-ФЗ · запрос котировок", etp="РТС-тендер")
C, b = omsk("0352500002526000089")
add(num="0352500002526000089", cat="Спецодежда", cust="БУ Омской обл. (Большереченское ДРСУ)", subj="Спецодежда и обувь зимняя", end="26.10.2026 05:00",
    nmck=498211.82, C=C, d=0.10, basis=b, law="44-ФЗ · запрос котировок", etp="РТС-тендер")
add(num="32616480814", cat="Арматура", cust="АО «Петербургский нефтяной терминал»", subj="Задвижки 30с41нж/30с941нж, 70 шт", end="26.10.2026 10:00",
    nmck=3163250.24, C=2066075 + 41322, d=0.15, basis="СибЗТА прайс 2026; 11 задвижек с электроприводом по 90 тыс. (если без приводов — себес ~1,1 млн)",
    law="223-ФЗ · конкурентный отбор", etp="ZakazRF (АГЗРТ)")
add(num="32616346943", cat="Арматура", cust="ПАО «Камчатскэнерго» (РусГидро)", subj="Арматура среднего/низкого давления, 315 позиций", end="29.10.2026 10:00",
    nmck=101102908, C=49931509, d=0.15, days=120, basis="315 позиций по ценам 2026 (252 с ценой поставщика, 63 оценки), торг −10%, доставка 55 ₽/кг",
    law="223-ФЗ · конкурс в эл. форме", etp="Росэлторг (РусГидро)", risk="аккредитация РусГидро, реестр МПТ по каждой позиции, БГ на аванс, обеспечение 1 млн")
add(num="32616450774", cat="Трубы ППУ", cust="МУП «Уфимские инженерные сети» (банкрот)", subj="Отводы в ППУ-ПЭ, 9 шт", end="20.10.2026 10:00",
    nmck=639848.52, C=478894, d=0.10, basis="4 завода, медиана цен 2026 + доставка 50 тыс.", law="223-ФЗ · запрос котировок", etp="РЭСТ",
    risk="СЕГОДНЯ до 10:00 МСК; заказчик в конкурсном производстве — отгрузка только после аванса 50%")
add(num="B1510261835045", cat="Форма", cust="АНО «Таврида.Арт», Судак", subj="Форменная одежда СБ: 8 позиций, 145 комплектов (по ТЗ)", end="30.10.2026 12:00",
    nmck=4169750, C=2925900, d=0.10, basis="по ТЗ: костюм 3 350, кофта 2 300, куртка 3 500, шапка 550, полуботинки 2 900, жилет 4 500, перчатки 1 500, снуд 420 ₽ + доставка в Судак 120 тыс.",
    law="Коммерческий тендер", etp="Росэлторг (коммерч.)", href="https://www.roseltorg.ru/procedures/search?query_field=B1510261835045",
    risk="поставка 10 дней — нужен товар в наличии; оплата 30 дн. после приёмки; цены кофты, ботинок, жилета, перчаток — подтвердить КП")
# оценки без спецификации
EST = [dict(num="386-076", cat="Обувь", cust="ООО «Союзспецодежда»", subj="Ботинки с высокими берцами, 4 000 пар", end="26.10.2026 17:40", nmck=10800000,
            C=9600000, d=0.0, basis="оценка: 2 400 ₽/пара (Спецобъединение/Тракт 2026); НМЦК нет — сумма = наше КП", law="Коммерческий тендер", etp="Bidzaar",
            href="https://bidzaar.com/"),
       dict(num="8004", cat="Арматура", cust="ООО «РМЗ», Салават", subj="Затворы дисковые", end="26.10.2026 13:00", nmck=10474222, C=8169893 + 245097, d=0.10,
            basis="оценка по категории (~20% маржа), спецификация на площадке", law="Коммерческий · конкурентный отбор", etp="ЭТП «Торги РФ»", href="https://торги.рф/"),
       dict(num="8002", cat="Арматура", cust="ООО «РМЗ», Салават", subj="Запорная арматура", end="26.10.2026 13:00", nmck=4248523, C=3313848 + 99415, d=0.10,
            basis="оценка по категории", law="Коммерческий · конкурентный отбор", etp="ЭТП «Торги РФ»", href="https://торги.рф/")]

thin = Side(style="thin", color="BFBFBF"); M = '#,##0" ₽"'; PCT = "0%"


def head(ws):
    for c in ws[ws.max_row]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")


wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Заходим"
ws.append(["Куда заходим — выгодные лоты без аукционов, маржа перепроверена 20.10.2026"]); ws["A1"].font = Font(bold=True, size=13)
ws.append(["Себестоимость — с НДС поставщика, с доставкой до заказчика; прибыль — после УСН 15%, с 0,8% расходов и стоимостью денег 20% годовых. Время подачи — МСК."])
H = ["№", "Подача до (МСК)", "Категория", "Номер", "Заказчик", "Предмет", "Способ", "Площадка", "НМЦК, ₽", "Себестоимость, ₽", "Маржа при НМЦК",
     "Ожид. снижение", "Цена заявки, ₽", "Прибыль после УСН, ₽", "В ноль при снижении", "Как посчитана себестоимость", "Риски", "Ссылка ЕИС", "Ссылка площадки"]
ws.append(H); head(ws)


def put(r, i, est=False):
    days = r.get("days", 60); P = r["nmck"] * (1 - r["d"]); net = model(P, r["C"], days)
    lk = L.get(r["num"], {})
    href = r.get("href") or lk.get("href") or ""
    etpu = lk.get("etp_url") or r.get("href") or ""
    if etpu and not etpu.startswith("http"): etpu = "https://" + etpu
    ws.append([i, r["end"], r["cat"], r["num"], r["cust"], r["subj"], r["law"], r["etp"], r["nmck"], r["C"], 1 - r["C"] / r["nmck"], r["d"], P, net, be(r["nmck"], r["C"], days),
               r["basis"], r.get("risk", "оценка — нужна спецификация" if est else ""), href if "zakupki" in href else "", etpu])
    R = ws.max_row; ws.cell(R, 4).number_format = "@"
    for j in (9, 10, 13, 14): ws.cell(R, j).number_format = M
    for j in (11, 12, 15): ws.cell(R, j).number_format = PCT
    for j in (18, 19):
        v = ws.cell(R, j).value
        if v: ws.cell(R, j).hyperlink = v; ws.cell(R, j).font = Font(color="0563C1", underline="single")
    if est:
        for c in ws[R][:17]: c.fill = PatternFill("solid", fgColor="F2F2F2")
    elif net >= 100000 and 1 - r["C"] / r["nmck"] >= 0.3:
        ws.cell(R, 14).fill = PatternFill("solid", fgColor="C6EFCE")
    for c in ws[R]: c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
    return net


ROWS.sort(key=lambda r: -model(r["nmck"] * (1 - r["d"]), r["C"], r.get("days", 60)))
tot = sum(put(r, i) for i, r in enumerate(ROWS, 1))
ws.append(["", "", "", "", "ИТОГО по подтверждённым", "", "", "", sum(r["nmck"] for r in ROWS), sum(r["C"] for r in ROWS), None, None, None, tot])
for j in (9, 10, 14): ws.cell(ws.max_row, j).number_format = M
for c in ws[ws.max_row]: c.font = Font(bold=True)
ws.append([]); ws.append(["", "", "", "", "Перспективные — только оценка, нужна спецификация"]); ws.cell(ws.max_row, 5).font = Font(bold=True)
for i, r in enumerate(EST, len(ROWS) + 1): put(r, i, est=True)
for i, w in enumerate([4, 15, 11, 20, 30, 34, 22, 16, 14, 14, 9, 9, 14, 14, 10, 50, 34, 40, 40], 1):
    ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "E4"
tidy(wb, row_h=48)
wb.save("../GO_LIST_2010.xlsx")
for r in ROWS: print(r["num"], round(r["C"]), round(1 - r["C"] / r["nmck"], 3), round(model(r["nmck"] * (1 - r["d"]), r["C"], r.get("days", 60))))
print("итого", round(tot), "TC", TC, "Gorka", GORKA)
