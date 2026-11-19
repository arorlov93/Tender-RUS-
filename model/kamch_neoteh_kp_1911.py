"""КП ООО «ГИДРОПАС» → ООО «НЕОТЕХ» (участник с лучшей ценой по Камчатскэнерго 32616346943):
поставка всего объёма лота 0200-РЕМ ПРОД-2027-КамчЭн за 70 000 000 ₽ (УСН, без НДС).
Цены = наша заявка 84 518 812,74 × k, сумма ровно 70 млн. → ../KAMCHATKA_NEOTEH_KP/"""
import json, os
from docx import Document
from docx.shared import Pt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill

OUT = "../KAMCHATKA_NEOTEH_KP"; os.makedirs(OUT, exist_ok=True)
TOTAL = 70_000_000.00
P = json.load(open("data/kamch/forms/kp_prices_84518812.json"))
T = json.load(open("data/kamch/tz_verbatim.json"))
base = sum(P[f"{t['table']}/{t['n_table']}"] * float(t["qty"]) for t in T)
k = TOTAL / base
rows = []
for t in T:
    q = float(t["qty"]); u = round(P[f"{t['table']}/{t['n_table']}"] * k, 2)
    rows.append([t["sec"], t["table"], t["n_table"], t["Наименование"], t.get("ГОСТ/ТУ") or "", t.get("Ду, мм") or "", t.get("Давление, МПа") or "", t["unit"], q, u, round(u * q, 2)])
diff = round(TOTAL - sum(r[10] for r in rows), 2)  # копейки округления — в позицию с qty=1 или в последнюю
j = next((i for i, r in enumerate(rows) if r[8] == 1), len(rows) - 1)
rows[j][9] = round(rows[j][9] + diff / rows[j][8], 2); rows[j][10] = round(rows[j][9] * rows[j][8], 2)
tot = round(sum(r[10] for r in rows), 2); qty = sum(r[8] for r in rows)

wb = Workbook(); ws = wb.active; ws.title = "Спецификация"
ws.append(["Приложение к коммерческому предложению ООО «ГИДРОПАС» от 19.11.2026 — поставка трубопроводной арматуры (лот № 0200-РЕМ ПРОД-2027-КамчЭн, ПАО «Камчатскэнерго»)"])
ws["A1"].font = Font(bold=True)
ws.append(["Раздел / грузополучатель", "Таблица ТЗ", "№ в табл.", "Наименование", "ГОСТ/ТУ", "Ду, мм", "Р, МПа", "Ед.", "Кол-во", "Цена за ед., руб. (НДС не облагается)", "Сумма, руб."])
for c in ws[2]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="0B2545"); c.alignment = Alignment(wrap_text=True, vertical="top")
for r in rows: ws.append(r)
ws.append(["ИТОГО", "", "", f"{len(rows)} позиций", "", "", "", "", qty, "", tot])
for c in ws[ws.max_row]: c.font = Font(bold=True)
ws.append(["НДС не облагается: ООО «ГИДРОПАС» применяет УСН (ст. 346.11 НК РФ). Цены включают доставку до грузополучателей ПАО «Камчатскэнерго» (г. Петропавловск-Камчатский)."])
for i, w in enumerate([34, 9, 8, 48, 22, 8, 8, 6, 9, 16, 16], 1): ws.column_dimensions[chr(64 + i)].width = w
for row in ws.iter_rows(min_row=3, max_row=ws.max_row - 1):
    for c in row[3:5]: c.alignment = Alignment(wrap_text=True, vertical="top")
    row[9].number_format = row[10].number_format = "# ##0.00"
ws.freeze_panes = "E3"
wb.save(f"{OUT}/Спецификация_ГИДРОПАС_для_НЕОТЕХ.xlsx")

fmt = lambda x: f"{x:,.0f}".replace(",", " ")
S = "Коммерческое предложение: арматура по лоту 0200-РЕМ ПРОД-2027-КамчЭн (ПАО «Камчатскэнерго»)"
BODY = f"""Добрый день!

ООО «ГИДРОПАС» предлагает поставку полного объёма продукции по конкурсу 32616346943 «Поставка трубопроводной арматуры среднего, низкого давления и механизмов управления для нужд ПАО «Камчатскэнерго»» (лот № 0200-РЕМ ПРОД-2027-КамчЭн).

Наше предложение:
• объём — все {len(rows)} позиций спецификации лота ({fmt(qty)} ед.) в точном соответствии с техническими требованиями заказчика;
• продукция российского производства с записями в реестре Минпромторга (ПП № 719) — под национальный режим закупки;
• общая стоимость — {fmt(TOTAL)} руб., НДС не облагается (УСН), включая доставку до грузополучателей в г. Петропавловске-Камчатском;
• сроки поставки, порядок оплаты и документы — по договорённости, готовы работать на условиях вашего договора с заказчиком;
• комплект документов: паспорта, сертификаты/декларации соответствия, выписки из реестра ПП 719.

Поштучная спецификация с ценами — во вложении. Предложение действует 30 дней.

Готовы обсудить детали и выслать проект договора.

С уважением,
Александр Орлов
ООО «ГИДРОПАС»
ИНН 9729231965, КПП 772601001
117105, г. Москва, Варшавское шоссе, д. 1А, помещ. 28/2
+7 (925) 878-06-78 | info@gidropas.ru"""

d = Document(); st = d.styles["Normal"]; st.font.name = "Times New Roman"; st.font.size = Pt(12)
d.add_paragraph("ООО «ГИДРОПАС»\nИНН 9729231965, КПП 772601001\n117105, г. Москва, Варшавское шоссе, д. 1А, помещ. 28/2\n+7 (925) 878-06-78, info@gidropas.ru")
d.add_paragraph("Исх. № ____ от 19.11.2026\n\nООО «НЕОТЕХ»\nИНН 3812084296\n664043, г. Иркутск, ул. Воронежская, д. 8\nmail@neoteh.info")
d.add_heading("Коммерческое предложение", 1)
for p in BODY.split("\n\n")[1:]: d.add_paragraph(p)
d.save(f"{OUT}/КП_ГИДРОПАС_НЕОТЕХ_19.11.2026.docx")
open(f"{OUT}/ПИСЬМО_для_отправки.txt", "w").write(f"Кому: mail@neoteh.info\nТема: {S}\nВложения: КП_ГИДРОПАС_НЕОТЕХ_19.11.2026.docx, Спецификация_ГИДРОПАС_для_НЕОТЕХ.xlsx\n\n{BODY}\n")
print("k =", round(k, 4), "итого", tot, "позиций", len(rows), "ед.", qty)
