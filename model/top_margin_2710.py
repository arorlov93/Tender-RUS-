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


import os
HIGH = {}
for f in ("high_A", "high_B1", "high_B2"):
    if os.path.exists(f"data/2610/{f}.json"):
        for x in json.load(open(f"data/2610/{f}.json")): HIGH[x["num"]] = x
R = []
for k in ("A1", "A2", "B", "C", "D"):
    for x0 in json.load(open(f"data/2610/deep_{k}.json")):
        x = HIGH.get(x0["num"], x0); x["_re"] = x0["num"] in HIGH; x["_m0"] = pct(x0.get("margin_pct"))
        m = pct(x.get("margin_pct")); e = end_of(x.get("end") or T.get(x["num"], {}).get("end"))
        if m is None or m <= 0 or str(x.get("decision")).startswith("Нет") or (e and e < NOW): continue
        R.append(dict(m=m, dec=x.get("decision"), end=e, grp=GRP[k], num=x["num"], name=str(x.get("name")), cust=str(x.get("cust") or ""), nmck=x.get("nmck"),
                      cost=(x.get("cost_total") or 0) + (x.get("logistics") or 0) or None, net=x.get("net_expected"), conf=x.get("confidence"), note=(("ПЕРЕСЧИТАНО 27.10: было " + (f'{x["_m0"]:.0%}' if x["_m0"] is not None else "—") + ". Не хватает: " + str(x.get("missing") or "")) if x["_re"] else str(x.get("risks") or ""))[:300], re=x["_re"],
                      link=T.get(x["num"], {}).get("link") or ""))
# ранее посчитанные лоты, по которым подача ещё открыта
R += [dict(m=0.458, dec="Участвовать", end=dt.datetime(2026, 11, 2, 10, 0), grp="Арматура", num="32616346943", name="Камчатскэнерго: трубопроводная арматура СиНД (315 поз.)",
           cust="АО «Камчатскэнерго» (РусГидро)", nmck=101102908, cost=54810725, net=22809966, conf="high", re=False,
           note="полный построчный расчёт и счета поставщиков (KAMCHATKA_STATUS_2610); 163 позиции ещё не котированы", link="https://zakupki.gov.ru/epz/order/notice/notice223/common-info.html?regNumber=32616346943"),
      dict(m=0.366, dec="Участвовать (опыт ≥30% НМЦ подтверждён 28.10)", end=dt.datetime(2026, 11, 5, 10, 0), grp="Арматура", num="11202-РЕМ ПРОД-2027-СЭ", name="Сахалинэнерго, лот 2: арматура СиНД для ЮС ТЭЦ-1 (87 поз.)",
           cust="ПАО «Сахалинэнерго» (РусГидро)", nmck=16153216.67, cost=10239786, net=2982228, conf="med", re=True,
           note="КП 28.10 (ПЭК, Доминик, АРМАТЭК ЗИП, ФОБОС ФБ39) учтены; ПЕРЕСЧИТАНО 27.10: все 87 позиций + контейнер до Южно-Сахалинска; 14% НМЦ — оценка; ~2 млн прибыли даёт п.53 (задвижки Ду500 с ЭП) — нужны КП Авангард/Тулаэлектропривод; опыт есть (подтверждено 28.10); запросы на заводы — SAKHALIN_ZAPROS", link="https://rushydro.roseltorg.ru"),
      dict(m=0.141, dec="Осторожно: КП АРМАТЭК получено, поз. 38 в минусе", end=dt.datetime(2026, 11, 5, 10, 0), grp="Арматура", num="11201-РЕМ ПРОД-2027-СЭ", name="Сахалинэнерго, лот 1: арматура СиНД для Сахалинской ГРЭС (83 поз.)",
           cust="ПАО «Сахалинэнерго» (РусГидро)", nmck=12269140, cost=10535025, net=-384272, conf="med", re=True,
           note="КП 28.10 (АРМАТЭК, ПЭК, ВФЗ, Лидер) учтены; АРМАТЭК поз.38 дороже НМЦ на 40%; ПЕРЕСЧИТАНО 27.10: все 83 позиции + доставка; 39% НМЦ — оценка (затворы АРМАТЭК АМ 2,4 млн — цены нет); 16 убыточных строк; при снижении >7% — минус", link="https://rushydro.roseltorg.ru"),
      dict(m=0.175, dec="Осторожно", end=dt.datetime(2026, 10, 30, 12, 0), grp="Одежда/СИЗ/соль", num="B1510261835045", name="Таврида.Арт: форменная одежда службы безопасности",
           cust="АНО «Таврида.Арт»", nmck=4169750, cost=3440900, net=319215, conf="high", re=False, note="построчно сверено (TAVRIDA_FINAL_2010); идти при подтверждённом наличии", link="")]
# сроки договоров: 44-ФЗ — из извещений ЕИС (gosplan), 223-ФЗ — из планов-графиков (contractEndDate), РусГидро — из документации
TERMS = {"0112200000826004798": "до 31.12.2026", "0848300064126000376": "49 кал. дн. с даты контракта", "0318300028426000179": "до 17.12.2026",
         "0372100001026000499": "01.01–31.12.2027, по заявкам до исчерпания суммы", "0137100001526000235": "до 31.12.2026", "0327300018526000001": "до 30.11.2026",
         "7600": "до 09.12.2026 (поставка 10 раб. дн.)", "8002": "до 31.12.2026", "0860200000826008623": "до 22.12.2026", "0829500001126007015": "до 31.12.2026",
         "0334200012526000832": "до 31.12.2026", "0322200005226000561": "до 10.08.2027 (поставки этапами до 10.12.2026 и 15.02.2027)",
         "0322200005226000562": "до 10.08.2027 (поставки этапами до 10.12.2026 и 15.02.2027)", "0372200087626000024": "до 31.12.2026",
         "32616346943": "поставка 120 кал. дн. с даты договора", "11202-РЕМ ПРОД-2027-СЭ": "поставка 220 кал. дн. с даты договора", "11201-РЕМ ПРОД-2027-СЭ": "поставка 230 кал. дн. с даты договора", "B1510261835045": "поставка 10 кал. дн. с даты договора, оплата 30 дн. после приёмки"}
if os.path.exists("data/2610/contract_ends.json"):
    for n, v in json.load(open("data/2610/contract_ends.json")).items():
        if v and v.get("end"):
            yrs = [y for y, sm in v.get("years", []) if float(sm or 0) > 0]
            TERMS[n] = f'договор до {dt.datetime.strptime(v["end"], "%Y-%m-%d"):%d.%m.%Y}' + (f' (оплата/поставка: {", ".join(yrs)} гг.)' if len(yrs) > 1 else (f" (весь объём в {yrs[0]} г.)" if yrs else "")) + " — план-график"
for r in R: r["term"] = TERMS.get(r["num"], "в документации (ЕИС/площадка недоступны)")
for r in R: r["gap"] = (r["nmck"] - r["cost"]) if r.get("nmck") and r.get("cost") else ((r["nmck"] or 0) * r["m"])
R = [r for r in R if r["gap"] > 1_000_000]  # только где НМЦК − себестоимость > 1 млн ₽
R.sort(key=lambda r: -r["m"])

thin = Side(style="thin", color="BFBFBF"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin); M = '#,##0" ₽"'
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "По марже"
ws.append([f"Открытые закупки без аукционов с разницей НМЦК − себестоимость больше 1 млн ₽ — по марже от большей к меньшей (на {NOW:%d.%m.%Y})"]); ws["A1"].font = Font(bold=True, size=13)
ws.append(["Маржа = 1 − (себестоимость + логистика) / НМЦК. Прибыль — после УСН 15% при ожидаемом снижении. Надёжность: high — построчно по спецификации, med — частично, low — оценка."])
H = ["№", "Маржа", "Решение", "Подача до (МСК)", "Срок договора / поставки", "Группа", "Номер", "Предмет", "Заказчик", "НМЦК, ₽", "Себестоимость + логистика, ₽", "Разница НМЦК − себестоимость, ₽", "Ожидаемая прибыль, ₽", "Надёжность", "Пересчитано 27.10", "Комментарий", "Ссылка"]
ws.append(H)
for c in ws[3]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")
for i, r in enumerate(R, 1):
    ws.append([i, r["m"], r["dec"], r["end"].strftime("%d.%m.%Y %H:%M") if r["end"] else "", r["term"], r["grp"], r["num"], r["name"][:90], r["cust"][:50], r["nmck"], r["cost"], r["gap"], r["net"], r["conf"], "да" if r.get("re") else "", r["note"], r["link"]])
    X = ws.max_row; ws.cell(X, 7).number_format = "@"; ws.cell(X, 2).number_format = "0%"
    for j in (10, 11, 12, 13): ws.cell(X, j).number_format = M
    ws.cell(X, 2).fill = PatternFill("solid", fgColor="C6EFCE" if r["m"] >= 0.3 else "E2EFDA" if r["m"] >= 0.2 else "FFEB9C" if r["m"] >= 0.1 else "F8CBAD")
    if r.get("re"): ws.cell(X, 15).fill = PatternFill("solid", fgColor="DDEBF7")
    if str(r["conf"]) == "low": ws.cell(X, 14).fill = PatternFill("solid", fgColor="FCE4D6")
    if r["link"]: ws.cell(X, 17).hyperlink = r["link"]; ws.cell(X, 17).font = Font(color="0563C1", underline="single")
    for c in ws[X]: c.border = BOX
for i, w in enumerate([4, 8, 18, 15, 26, 16, 22, 44, 30, 15, 16, 16, 15, 10, 10, 70, 30], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "H4"; ws.auto_filter.ref = f"A3:Q{ws.max_row}"
tidy(wb, row_h=44)
wb.save("../TOP_MARGIN_2710.xlsx")
for r in R: print(f'{r["m"]:.0%}', r["end"], r["num"], r["dec"][:22], round(r["net"] or 0), r["conf"], r["name"][:50])
