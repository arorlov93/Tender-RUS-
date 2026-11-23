"""Nostra Trade (PL) — Coca-Cola / Fanta / Monster: себестоимость банки в Москве по ценам прайса и при скидке 5/10 %.
Расчёт как по Red Bull: банка + доставка фурой (3 000 € до Москвы) + пошлина 8 % (мин. 0,03 €/л) + акциз РФ 11 ₽/л (сахаросодерж.)
+ НДС 22 %. Курс 92,19 ₽/€. Фура ограничена весом ~23 т брутто. → ../NOSTRA_COLA_RF_2311.xlsx"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
EUR = 92.19; FREIGHT = 3000; DUTY = .08; EXC = 11.0; VAT = .22; MAXKG = 23000
# товар, л, € за банку (прайс 19.01), банок/паллет, кг брутто/банка, акциз?
G = [("Coca-Cola 330 мл slim", .33, .3794, 2640, .355, True), ("Coca-Cola 330 мл (обычная банка)", .33, .4205, 2376, .365, True),
     ("Coca-Cola Cherry 330 мл", .33, .4205, 2376, .365, True), ("Coca-Cola Vanilla 330 мл", .33, .4205, 2376, .365, True),
     ("Fanta Orange / Lemon / Exotic 330 мл", .33, .4205, 2376, .365, True), ("Fanta Strawberry Kiwi 330 мл", .33, .3906, 2376, .365, True),
     ("Monster Ultra Ruby Red Zero 500 мл", .50, .8204, 1728, .54, False)]
wb = Workbook(); ws = wb.active; ws.title = "Москва, ₽ за банку"
H = PatternFill("solid", fgColor="0B2545"); HF = Font(bold=True, color="FFFFFF"); B = PatternFill("solid", fgColor="E2EFDA")
ws.append(["Nostra Trade → Москва: себестоимость банки (₽), курс 92,19 ₽/€, фура 3 000 €"]); ws["A1"].font = Font(bold=True, size=12)
hd = ["Товар", "Сценарий", "Цена €/банка", "Паллет в фуре (по весу)", "Банок в фуре", "Банка, ₽", "Доставка, ₽", "Пошлина 8 %, ₽", "Акциз, ₽", "НДС 22 %, ₽",
      "ИТОГО ₽/банка", "ИТОГО ₽/лоток 24", "Фура: товар, €", "Фура: всё в Москве, ₽"]
ws.append(hd)
for c in ws[2]: c.fill = H; c.font = HF; c.alignment = Alignment(wrap_text=True, vertical="top")
for name, l, p, cpp, kg, exc in G:
    pal = min(33, int(MAXKG // (cpp * kg + 25))); cans = pal * cpp
    for lab, d in (("прайс", 0), ("скидка −5 %", .05), ("скидка −10 %", .10)):
        pe = p * (1 - d); fr = FREIGHT / cans
        duty = max(DUTY * (pe + fr), .03 * l); ex = EXC * l if exc else 0
        base = (pe + fr + duty) * EUR + ex; vat = base * VAT; tot = base + vat
        ws.append([name, lab, round(pe, 4), pal, cans, round(pe * EUR, 2), round(fr * EUR, 2), round(duty * EUR, 2), round(ex, 2), round(vat, 2),
                   round(tot, 2), round(tot * 24, 0) if l < .4 else round(tot * 12, 0), round(pe * cans), round(tot * cans)])
        if d == .10:
            for c in ws[ws.max_row]: c.fill = B
for col, w in zip("ABCDEFGHIJKLMN", [34, 13, 10, 10, 10, 9, 9, 10, 8, 9, 11, 11, 12, 14]): ws.column_dimensions[col].width = w
for row in ws.iter_rows(min_row=3):
    for c in row[12:]: c.number_format = "# ##0"
ws.freeze_panes = "C3"
ws.append([]); notes = [
 "Цены Nostra — прайс 19.01 (склад Młyny, PL), вероятно EXW; условия и минимальная партия — уточнить.",
 "Фура ограничена весом: 330 мл — 24–26 паллет (не 33), Monster 500 мл — 24 паллеты; доставка 3 000 € распределена на фактическое число банок.",
 "Пошлина 8 % от (цена + доставка), но не менее 0,03 €/л; акциз РФ 11 ₽/л на сахаросодержащие (Monster Zero — без акциза); НДС 22 % на всё.",
 "Не включены: таможенное оформление/брокер, «Честный знак», сертификация ЕАЭС, стикеровка, хранение — +1,5–3 ₽/банка.",
 "Риск: Coca-Cola и Fanta (код 2202) в перечне параллельного импорта (приказ Минпромторга 2701, ред. с 27.05.2026) не значатся — ввоз без согласия правообладателя = риск по товарному знаку (как с Red Bull). Проверить у юриста."]
for n in notes: ws.append([n])
wb.save("../NOSTRA_COLA_RF_2311.xlsx")
for r in ws.iter_rows(min_row=3, max_row=3 + 3 * len(G) - 1, values_only=True): print(r[0][:22], r[1], r[3], r[10])
