"""Открытые закупки (без аукционов, подача не прошла) — по марже от большей к меньшей: разборы 26.10 (data/2610/deep_*.json) + Камчатка, Таврида, Сахалинэнерго. → TOP_MARGIN_2710.xlsx"""
import json, re, datetime as dt, openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side
from tidy import tidy

NOW = dt.datetime(2026, 10, 27, 12, 0)
T = {x["num"]: x for x in json.load(open("data/triage_2610.json"))}
pct = lambda v: None if v is None else (v / 100 if abs(v) > 1.5 else v)
GRP = {"A1": "Арматура/ППУ", "A2": "Арматура", "B": "Одежда/СИЗ/соль", "C": "Пробирки", "D": "Видеонаблюдение/IT"}


def end_of(s):
    s = str(s or "")
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})", s) or re.search(r"(\d{2})\.(\d{2})\.(\d{4})\D+(\d{1,2}):(\d{2})", s)
    if not m: return None
    g = m.groups()
    return dt.datetime(int(g[0]), int(g[1]), int(g[2]), int(g[3]), int(g[4])) if len(g[0]) == 4 else dt.datetime(int(g[2]), int(g[1]), int(g[0]), int(g[3]), int(g[4]))


R = []
for k in ("A1", "A2", "B", "C", "D"):
    for x in json.load(open(f"data/2610/deep_{k}.json")):
        m = pct(x.get("margin_pct")); e = end_of(x.get("end") or T.get(x["num"], {}).get("end"))
        if m is None or m <= 0 or str(x.get("decision")).startswith("Нет") or (e and e < NOW): continue
        R.append(dict(m=m, dec=x.get("decision"), end=e, grp=GRP[k], num=x["num"], name=str(x.get("name")), cust=str(x.get("cust") or ""), nmck=x.get("nmck"),
                      cost=(x.get("cost_total") or 0) + (x.get("logistics") or 0) or None, net=x.get("net_expected"), conf=x.get("confidence"), note=str(x.get("risks") or "")[:220],
                      link=T.get(x["num"], {}).get("link") or ""))
# ранее посчитанные лоты, по которым подача ещё открыта
R += [dict(m=0.458, dec="Участвовать", end=dt.datetime(2026, 11, 2, 10, 0), grp="Арматура", num="32616346943", name="Камчатскэнерго: трубопроводная арматура СиНД (315 поз.)",
           cust="АО «Камчатскэнерго» (РусГидро)", nmck=101102908, cost=54810725, net=22809966, conf="high",
           note="полный построчный расчёт и счета поставщиков (KAMCHATKA_STATUS_2610); 163 позиции ещё не котированы", link="https://zakupki.gov.ru/epz/order/notice/notice223/common-info.html?regNumber=32616346943"),
      dict(m=0.441, dec="Участвовать — при опыте 4,85 млн", end=dt.datetime(2026, 11, 5, 10, 0), grp="Арматура", num="11202-РЕМ ПРОД-2027-СЭ", name="Сахалинэнерго, лот 2: арматура СиНД для ЮС ТЭЦ-1 (87 поз.)",
           cust="ПАО «Сахалинэнерго» (РусГидро)", nmck=16153216.67, cost=None, net=None, conf="med",
           note="маржа по 63 из 87 позиций (56% НМЦ); 24 позиции без цены; отборочное требование — опыт ≥ 30% НМЦ (SAKHALIN_11201_11202)", link="https://rushydro.roseltorg.ru"),
      dict(m=0.230, dec="Осторожно — при опыте 3,68 млн", end=dt.datetime(2026, 11, 5, 10, 0), grp="Арматура", num="11201-РЕМ ПРОД-2027-СЭ", name="Сахалинэнерго, лот 1: арматура СиНД для Сахалинской ГРЭС (83 поз.)",
           cust="ПАО «Сахалинэнерго» (РусГидро)", nmck=12269140, cost=None, net=None, conf="med",
           note="маржа по 54 из 83 позиций (47% НМЦ); затворы п.19–20 убыточны; опыт ≥ 30% НМЦ", link="https://rushydro.roseltorg.ru"),
      dict(m=0.175, dec="Осторожно", end=dt.datetime(2026, 10, 30, 12, 0), grp="Одежда/СИЗ/соль", num="B1510261835045", name="Таврида.Арт: форменная одежда службы безопасности",
           cust="АНО «Таврида.Арт»", nmck=4169750, cost=3440900, net=319215, conf="high", note="построчно сверено (TAVRIDA_FINAL_2010); идти при подтверждённом наличии", link="")]
R.sort(key=lambda r: -r["m"])

thin = Side(style="thin", color="BFBFBF"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin); M = '#,##0" ₽"'
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "По марже"
ws.append([f"Открытые закупки без аукционов — по марже от большей к меньшей (на {NOW:%d.%m.%Y})"]); ws["A1"].font = Font(bold=True, size=13)
ws.append(["Маржа = 1 − (себестоимость + логистика) / НМЦК. Прибыль — после УСН 15% при ожидаемом снижении. Надёжность: high — построчно по спецификации, med — частично, low — оценка."])
H = ["№", "Маржа", "Решение", "Подача до (МСК)", "Группа", "Номер", "Предмет", "Заказчик", "НМЦК, ₽", "Себестоимость + логистика, ₽", "Ожидаемая прибыль, ₽", "Надёжность", "Комментарий", "Ссылка"]
ws.append(H)
for c in ws[3]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")
for i, r in enumerate(R, 1):
    ws.append([i, r["m"], r["dec"], r["end"].strftime("%d.%m.%Y %H:%M") if r["end"] else "", r["grp"], r["num"], r["name"][:90], r["cust"][:50], r["nmck"], r["cost"], r["net"], r["conf"], r["note"], r["link"]])
    X = ws.max_row; ws.cell(X, 6).number_format = "@"; ws.cell(X, 2).number_format = "0%"
    for j in (9, 10, 11): ws.cell(X, j).number_format = M
    ws.cell(X, 2).fill = PatternFill("solid", fgColor="C6EFCE" if r["m"] >= 0.3 else "E2EFDA" if r["m"] >= 0.2 else "FFEB9C" if r["m"] >= 0.1 else "F8CBAD")
    if str(r["conf"]) == "low": ws.cell(X, 12).fill = PatternFill("solid", fgColor="FCE4D6")
    if r["link"]: ws.cell(X, 14).hyperlink = r["link"]; ws.cell(X, 14).font = Font(color="0563C1", underline="single")
    for c in ws[X]: c.border = BOX
for i, w in enumerate([4, 8, 18, 15, 16, 22, 44, 30, 15, 16, 15, 10, 60, 30], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "G4"; ws.auto_filter.ref = f"A3:N{ws.max_row}"
tidy(wb, row_h=44)
wb.save("../TOP_MARGIN_2710.xlsx")
for r in R: print(f'{r["m"]:.0%}', r["end"], r["num"], r["dec"][:22], round(r["net"] or 0), r["conf"], r["name"][:50])
