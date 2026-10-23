"""ППУ Уфа (договор по 32616450774, цена 595 000 ₽): заводы-производители отводов, цены 2026, доставка до Уфы, прибыль по сценариям, кому слать запрос КП."""
import json, openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side
from tidy import tidy

PRICE = 595000
QTY = {"108": 3, "377": 2, "426": 3, "530": 1}
C = json.load(open("data/ppu_ufa/ppu_center.json")); U = json.load(open("data/ppu_ufa/ppu_ural.json")); FR = json.load(open("data/ppu_ufa/freight.json"))
thin = Side(style="thin", color="BFBFBF"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin); M = '#,##0" ₽"'


def head(ws):
    for c in ws[ws.max_row]:
        if c.value is not None: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")


def model(C_, L):  # как в ppu_ufa.py: аванс 50%, деньги 20% годовых 60 дн, 0,8% ЭП/банк, УСН 15%
    pre = PRICE - C_ - L - 0.008 * PRICE - (C_ + L) * 0.5 * 0.20 * 60 / 365
    return pre - 0.15 * max(pre, 0)


kit = lambda p: sum(p[k] * q for k, q in QTY.items())
# сценарии: цены за шт с НДС с сайтов (ПКФ ТЕПЛО — из data/ppu_prices_1910.json, тип 1, стенка 4,5 для 108)
SC = [
    ("НЗТИ (Н. Новгород) 377/426/530 + ПКФ ТЕПЛО 108х4,5", {"108": 6907.4, "377": 44743, "426": 54422, "530": 82733}, 45000,
     "оценка: сборный груз Н. Новгород → Уфа ≈ как из Ярославля (42,5–45,9 тыс. по Байкалу); две отгрузки — +≈15 тыс., если 108 отдельно", "1nzti.ru, ppu76.ru"),
    ("Всё у НЗТИ (108 — уточнить 4,5 мм, по сайту только 4,0 и 5,0)", {"108": 7424, "377": 44743, "426": 54422, "530": 82733}, 45000,
     "оценка, как выше; одна отгрузка", "1nzti.ru"),
    ("Всё у ПКФ ТЕПЛО (Ярославль) — текущий расчёт", {"108": 6907.4, "377": 53598.4, "426": 55110.3, "530": 91102.7}, 45917,
     "Байкал Сервис, сборный 6 м³ до двери, 5 сут", "ppu76.ru"),
    ("Уральский завод (АЗПТ / УЗТИ / ЧЗИТ / ОЗТИ) — если КП не выше ПКФ ТЕПЛО", {"108": 6907.4, "377": 53598.4, "426": 55110.3, "530": 91102.7}, 14395,
     "отдельная машина 3 т Екатеринбург/Челябинск → Уфа, 1 день (maintransport.ru)", "цены — по запросу; здесь — уровень ПКФ ТЕПЛО"),
]
wb = openpyxl.Workbook(); sv = wb.active; sv.title = "Итог"
sv.append(["Отводы ППУ-ПЭ для МУП «Уфимские инженерные сети»: где купить и как привезти. Наша цена договора 595 000 ₽ с НДС"]); sv["A1"].font = Font(bold=True, size=12)
sv.append(["Цены — с сайтов заводов, октябрь 2026, с НДС, за 1 шт («от», не оферта). Все цены — до подтверждения КП: отвод крутоизогнутый ГОСТ 17375 (не сварной/не из ГОСТ 10704), стенка как в ТЗ, ОДК 3/4 провода."])
sv.append(["Вариант", "108×3", "377×2", "426×3", "530×1", "Комплект, ₽", "Доставка до Уфы, ₽", "Итого себестоимость, ₽", "Прибыль после УСН, ₽", "Маржа", "Доставка — откуда цифра", "Источник цен"]); head(sv)
for n, p, L, ls, src in SC:
    k = kit(p); sv.append([n, p["108"] * 3, p["377"] * 2, p["426"] * 3, p["530"], k, L, k + L, model(k, L), model(k, L) / PRICE, ls, src])
    for j in range(2, 10): sv.cell(sv.max_row, j).number_format = M
    sv.cell(sv.max_row, 10).number_format = "0.0%"
sv.append([])
sv.append(["Что важно для приёмки", "", "", "", "", "", "", "", "", "", "", ""]); head(sv)
for t in ["НЗТИ на карточках пишет «Отвод ППУ ГОСТ 10704 Ст 10-20» — это может быть отвод из электросварной трубы. ТЗ требует отвод по ГОСТ 17375/17380, спиральношовные нельзя. В КП — только крутоизогнутый ГОСТ 17375; цена может вырасти.",
          "108: в ТЗ стенка 4,5 мм. У НЗТИ на сайте 108×4,0 и 108×5,0 — 4,0 не подходит; 5,0 — отклонение от ТЗ, лучше 4,5 у ПКФ ТЕПЛО (108×4,5/180, 6 907 ₽) или под заказ у НЗТИ.",
          "Отвод 530 — 4 проводника ОДК (3, 6, 9, 12 ч), остальные — 3 (3, 12, 9 ч); концы 150 мм (108) и 210 мм (остальные). Подтвердить в КП по листу «ТЗ дословно» (PPU_UFA_SPEC_KP.xlsx).",
          "Разгрузка: отвод 530 ≈ 333 кг — у Байкала разгрузка по Уфе индивидуально, гидроборта нет. Нужен погрузчик заказчика или манипулятор — согласовать с МУП до отгрузки."]:
    sv.append([t]); sv.merge_cells(start_row=sv.max_row, start_column=1, end_row=sv.max_row, end_column=12)
for row in sv.iter_rows(min_row=3):
    for c in row:
        if c.value is not None: c.border = BOX
for i, w in enumerate([44, 11, 11, 11, 11, 13, 13, 14, 14, 8, 44, 22], 1): sv.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

pr = wb.create_sheet("Цены заводов")
pr.append(["Компания", "Город", "Производитель?", "Позиция", "Цена за шт, ₽", "НДС", "Наличие", "Ссылка", "Примечание"]); head(pr)
for x in sorted([x for x in C + U if isinstance(x["price"], (int, float))], key=lambda x: (x["item"], x["price"])):
    pr.append([x["company"], x.get("plant_city"), x["is_manufacturer"][:40], x["item"], x["price"], x["vat"], str(x.get("in_stock") or "")[:20], x["url"], str(x.get("note") or "")[:200]])
    pr.cell(pr.max_row, 5).number_format = M
    if x["url"]: pr.cell(pr.max_row, 8).hyperlink = x["url"]; pr.cell(pr.max_row, 8).font = Font(color="0563C1", underline="single")
    for c in pr[pr.max_row]: c.border = BOX
for i, w in enumerate([34, 16, 16, 8, 12, 12, 12, 40, 70], 1): pr.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
pr.freeze_panes = "B2"; pr.auto_filter.ref = f"A1:I{pr.max_row}"

kp = wb.create_sheet("Запросить КП")
kp.append(["№", "Завод", "Город", "До Уфы, км", "Производитель?", "Телефон", "E-mail", "Сайт", "Где взят контакт", "Как проверен", "Почему в списке"]); head(kp)
seen, PRIO = set(), ["Нижегородский", "СМИТ", "Арамиль", "Уральский завод трубной", "Полевской", "Челябинский завод изоляции", "Оренбургский", "Уфимский трубный", "АО «ТИК»", "МосФлоуЛайн", "Теплосеть", "ЛТР"]
WHY = {"Нижегородский": "самые низкие цены с НДС на 377/426/530, в наличии", "СМИТ": "крупный завод, отводы на ГОСТ 17375; цены только по запросу",
       "Арамиль": "Урал — доставка 14 тыс. отдельной машиной", "Уральский завод трубной": "Урал — доставка 14 тыс.", "Полевской": "Урал — доставка 14 тыс.",
       "Челябинский завод изоляции": "Урал — доставка 14 тыс.", "Оренбургский": "ближе всех, доставка 14 тыс.", "Уфимский трубный": "в самой Уфе — без доставки; уточнить, работает ли цех",
       "АО «ТИК»": "производитель, Серпухов", "МосФлоуЛайн": "крупный производитель, Москва", "Теплосеть": "производитель (Киров/Саратов)", "ЛТР": "низкие цены на сайте — проверить стенку и тип отвода"}
n = 0
for key in PRIO:
    for x in C + U:
        if key in x["company"] and x["company"] not in seen:
            seen.add(x["company"]); n += 1
            kp.append([n, x["company"], x.get("plant_city"), x.get("distance_to_ufa_km"), x["is_manufacturer"][:30], x.get("phone"), x.get("email"), x.get("site"),
                       x.get("contact_url"), str(x.get("contact_check") or "")[:160], WHY[key]])
            for c in kp[kp.max_row]: c.border = BOX
for i, w in enumerate([4, 34, 16, 9, 14, 30, 26, 18, 30, 50, 34], 1): kp.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

fr = wb.create_sheet("Доставка")
fr.append(["Откуда", "Способ", "Стоимость, ₽", "Срок", "Как получено", "Ссылка"]); head(fr)
for x in FR:
    if x["price"] is None: continue
    fr.append([x["from"], x["mode"][:90], x["price"], x.get("days"), str(x.get("evidence") or "")[:160], x.get("url")])
    fr.cell(fr.max_row, 3).number_format = M
    for c in fr[fr.max_row]: c.border = BOX
for i, w in enumerate([30, 50, 12, 14, 60, 40], 1): fr.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
tidy(wb, row_h=40)
wb.save("../PPU_UFA_FACTORIES_2310.xlsx")
for n_, p, L, *_ in SC: print(n_[:45], round(kit(p)), L, round(model(kit(p), L)))
