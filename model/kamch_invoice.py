"""Камчатскэнерго 32616346943: сверка счёта поставщика (ДН.РУ № 819892 от 22.10.2026, задвижки) с нашей проверенной ценой по листу «Задвижки»."""
import json, re, openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side
from tidy import tidy

INV = json.load(open("data/kamch/kp/dnru_819892.json"))
I = {x["z"]: x for x in INV["items"]}
src = openpyxl.load_workbook("../KAMCHATKA_BY_GROUP_KP.xlsx")["Задвижки"]
POS = [r for r in src.iter_rows(values_only=True) if r[0] and str(r[0]).isdigit() and r[1] and "/" in str(r[1])]
TORG, FREIGHT = 0.10, 55  # как в kamch_verify: наша себестоимость = цена × 0,9 + доставка (одинакова в обоих вариантах)
COST_KAM = 54810725
thin = Side(style="thin", color="BFBFBF"); M = '#,##0" ₽"'; P = "+0%;-0%;0%"
F = {"g": "C6EFCE", "y": "FFEB9C", "r": "F8CBAD", "b": "DDEBF7"}


def head(ws):
    for c in ws[ws.max_row]:
        if c.value is not None: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")


rows = []
for r in POS:
    z, tt, name, dn, pn, mat, q, our = int(r[0]), r[1], r[2].split("·", 1)[1].strip(), r[3], r[4], str(r[5] or ""), float(r[7]), float(r[8])
    x = I.get(z)
    d = dict(z=z, tt=tt, name=name, q=q, mat=mat, our=our, inv=None, dev=None, srok="", art="", offer="", note="", col="b")
    if not x:
        d.update(note="Нет в счёте — ДН.РУ не котировал (запросить СИБЗТА / ПТПА / Алексин)")
        rows.append(d); continue
    d.update(inv=x["price_vat"], dev=x["price_vat"] / our - 1, srok=x["srok"], art=x["art"], offer=f'{x["name"]}; корпус {x["body"]}')
    notes = []
    if "30нж" in name and ("CF8" in x["body"] or "304" in x["body"]) and "12Х18Н9ТЛ" in mat.upper().replace(" ", ""):
        notes.append("Корпус CF8/AISI304 ≠ 12Х18Н9ТЛ из ТЗ — нужно согласие заказчика на эквивалент, иначе отклонят")
    if "КОФ" in name.upper():
        notes.append("В ТЗ «с КОФ» — в счёте задвижка без ответных фланцев, КОФ докупать")
    if d["dev"] > 0.5:
        notes.append("Дороже рынка в 2–4 раза — у ДН.РУ не брать, запросить СИБЗТА / ПТПА / БК-Арматура")
    elif d["dev"] > 0.1:
        notes.append("Дороже нашей цены — торговаться или брать у другого (БК-Арматура, СИБЗТА)")
    elif d["dev"] < -0.1:
        notes.append("Дешевле нашей цены")
    else:
        notes.append("Совпадает с нашей ценой (±10%)")
    d["col"] = "r" if d["dev"] > 0.5 else "y" if d["dev"] > 0.1 or len(notes) > 1 else "g"
    d["note"] = "; ".join(notes)
    rows.append(d)

have = [d for d in rows if d["inv"]]
s_our = sum(d["our"] * d["q"] for d in have); s_inv = sum(d["inv"] * d["q"] for d in have)
c_our = sum(d["our"] * (1 - TORG) * d["q"] for d in have)
c_best = sum(min(d["inv"], d["our"] * (1 - TORG)) * d["q"] for d in have)  # берём счёт ДН.РУ только там, где он дешевле нашей себестоимости
c_all = s_inv
c_safe = sum((min(d["inv"], d["our"] * (1 - TORG)) if "30нж" not in d["name"] else d["our"] * (1 - TORG)) * d["q"] for d in have)  # без нерж. CF8 (не по ТЗ)
miss = [d for d in rows if not d["inv"]]

wb = openpyxl.Workbook(); sv = wb.active; sv.title = "Итог"
sv.append([f'Камчатскэнерго 32616346943 — задвижки: {INV["invoice"]}, {INV["supplier"]}']); sv["A1"].font = Font(bold=True, size=12)
sv.append([f'Цены за 1 шт. с НДС 22%, счёт на 1 шт. каждой позиции, итого ' + f'{INV["total_vat"]:,.0f}'.replace(",", " ") + f' ₽. Счёт действителен {INV["valid"]}; самовывоз со склада в Москве, доставка до Камчатки — наша.'])
sv.append(["Показатель", "Значение", "Комментарий"]); head(sv)
for a, b, c, f in [
    ("Позиций «Задвижки» в нашей таблице", len(rows), f"{int(sum(d['q'] for d in rows))} шт.", "0"),
    ("Из них есть в счёте ДН.РУ", len(have), f"{int(sum(d['q'] for d in have))} шт.", "0"),
    ("Нет в счёте", len(miss), "крупные / с электроприводом / Ру63 — запросить у заводов", "0"),
    ("Наша цена по этим позициям × кол-во", s_our, "проверенная рыночная цена с НДС", M),
    ("Счёт ДН.РУ × кол-во", s_inv, f"{s_inv / s_our - 1:+.0%} к нашей цене", M),
    ("Наша себестоимость (цена −10%)", c_our, "как заложено в расчёте Камчатки", M),
    ("Себестоимость, если брать всё у ДН.РУ", c_all, f"{c_all - c_our:+,.0f} ₽ к расчёту".replace(",", " "), M),
    ("Себестоимость, если ДН.РУ — только где дешевле", c_best, f"{c_best - c_our:+,.0f} ₽ к расчёту".replace(",", " "), M),
    ("То же, но без нерж. CF8 (не по ТЗ)", c_safe, f"{c_safe - c_our:+,.0f} ₽ к расчёту — реалистичный вариант".replace(",", " "), M),
    ("Себестоимость всего лота Камчатки было", COST_KAM, "", M),
    ("Себестоимость лота с лучшими ценами", COST_KAM + c_best - c_our, f"маржа при НМЦК почти не меняется — экономия {c_our - c_best:,.0f} ₽".replace(",", " "), M),
]:
    sv.append([a, b, c]); sv.cell(sv.max_row, 2).number_format = f
sv.append([])
sv.append(["Группа", "Позиций в счёте", "Наша цена × кол-во, ₽", "Счёт × кол-во, ₽", "Отклонение", "Вывод"]); head(sv)
for g, rx, concl in [("30нж41нж (нерж.)", r"30нж41нж", "в 2 раза дешевле, но корпус CF8/AISI304 вместо 12Х18Н9ТЛ — брать только при согласии заказчика на эквивалент"),
                     ("30с41нж (сталь Ру16)", r"30с41нж", "совпадает с нашей ценой — можно брать (Ду50 на 11% дороже — торговаться)"),
                     ("30с64нж (сталь Ру25)", r"30с64нж", "в 2–4 раза дороже рынка — у ДН.РУ не брать (СИБЗТА 28 444 ₽ за Ду100)"),
                     ("30ч6бр (чугун)", r"30ч6бр", "Ду50–150 дороже на 20–42%, Ду200–300 дешевле на 1–13% — брать только Ду200–300")]:
    gg = [d for d in have if re.search(rx, d["name"])]
    a, b = sum(d["our"] * d["q"] for d in gg), sum(d["inv"] * d["q"] for d in gg)
    sv.append([g, len(gg), a, b, b / a - 1 if a else None, concl])
    for j in (3, 4): sv.cell(sv.max_row, j).number_format = M
    sv.cell(sv.max_row, 5).number_format = P
sv.append([])
sv.append(["Что уточнить у ДН.РУ", "", ""]); head(sv)
for t in ["Страна производства DN.ru GVWR/GVP и номер записи в реестре Минпромторга (ПП 719) — для нацрежима ПП 1875",
          "Скидка за объём: счёт на 1 шт., нам нужно 298 шт. задвижек",
          "Цена на 30с64нж — почему втрое выше рынка; есть ли другой бренд (Техгидромаш / LD)",
          "30нж41нж в 12Х18Н9ТЛ (исп. 1 САЗ) — есть ли; если нет, письмо заказчику об эквиваленте CF8",
          "Курсовая оговорка: счёт пересматривается при изменении курса > 5%, действует 2 дня — до подачи зафиксировать цену в КП"]:
    sv.append([t]); sv.merge_cells(start_row=sv.max_row, start_column=1, end_row=sv.max_row, end_column=6)
for row in sv.iter_rows(min_row=3):
    for c in row:
        if c.value is not None: c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
for i, w in enumerate([46, 20, 60, 18, 12, 70], 1): sv.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

ws = wb.create_sheet("Построчно")
ws.append(["№ по заявке", "ТТ / №", "Позиция ТЗ", "Материал по ТЗ", "Кол-во", "Наша цена, ₽/шт", "Счёт ДН.РУ, ₽/шт", "Отклонение", "Наша сумма, ₽", "Сумма по счёту, ₽",
           "Предложено в счёте", "Артикул", "Срок", "Вывод"]); head(ws)
for d in rows:
    ws.append([d["z"], d["tt"], d["name"], d["mat"][:40], d["q"], d["our"], d["inv"], d["dev"], d["our"] * d["q"], d["inv"] * d["q"] if d["inv"] else None,
               d["offer"], d["art"], d["srok"], d["note"]])
    X = ws.max_row
    for j in (6, 7, 9, 10): ws.cell(X, j).number_format = M
    ws.cell(X, 8).number_format = P
    for j in (7, 8, 14): ws.cell(X, j).fill = PatternFill("solid", fgColor=F[d["col"]])
ws.append(["", "", "ИТОГО (позиции из счёта)", "", sum(d["q"] for d in have), "", "", s_inv / s_our - 1, s_our, s_inv])
for j in (9, 10): ws.cell(ws.max_row, j).number_format = M
ws.cell(ws.max_row, 8).number_format = P
for c in ws[ws.max_row]: c.font = Font(bold=True)
for row in ws.iter_rows(min_row=1):
    for c in row: c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
for i, w in enumerate([8, 9, 40, 18, 7, 13, 13, 10, 14, 14, 44, 12, 16, 50], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "D2"; ws.auto_filter.ref = f"A1:N{ws.max_row}"
tidy(wb, row_h=44)
wb.save("../KAMCHATKA_KP_DNRU_819892.xlsx")
print(len(have), "в счёте;", len(miss), "нет | наша", round(s_our), "счёт", round(s_inv), f"{s_inv / s_our - 1:+.1%}", "| себес было", round(c_our), "всё ДН.РУ", round(c_all), "лучшее", round(c_best), "без CF8", round(c_safe))
