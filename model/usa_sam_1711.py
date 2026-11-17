"""SAM.gov 17.11.2026: разбор 20 закупок из выгрузки Claude in Chrome + пакет на Rock Salt M6700127Q1014.
→ ../ZAYAVKI_USA/SAM_RAZBOR_1711.xlsx, ../ZAYAVKI_USA/ROCK_SALT_M6700127Q1014/*"""
import csv, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from docx import Document
from docx.shared import Pt

OUT = "../ZAYAVKI_USA"; RS = f"{OUT}/ROCK_SALT_M6700127Q1014"
os.makedirs(RS, exist_ok=True)
H = PatternFill("solid", fgColor="0B2545"); HF = Font(bold=True, color="FFFFFF"); W = Alignment(wrap_text=True, vertical="top")
G = PatternFill("solid", fgColor="C6EFCE"); Y = PatternFill("solid", fgColor="FFEB9C"); R = PatternFill("solid", fgColor="F4CCCC")

# --- 1. Разбор всех 20 закупок -------------------------------------------------
VERDICT = {
 "M6700127Q1014": ("ИДЁМ", "Каменная соль 210 т/год ×3 года, full & open, best value. Товар американский, закупаем у дистрибьютора с пневмовозом. Прошлые победители — такие же перепродавцы (KGOC GA $28,7k, 10 участников; Cultivated SC $50,4k, 6 участников). Подача email до 23.11 16:00 ET."),
 "W911S227U0140": ("ПРОВЕРИТЬ", "Микрокристаллический воск #118P, Total Small Business, Yuma Proving Ground AZ, до 23.11 17:00 MST. ТЗ и количество только в PIEE (нужна регистрация вендора в PIEE). Поставщики воска в США: IGI, Sonneborn, Blended Waxes. Решение после PIEE."),
 "15B70027PR000089": ("ОТВЕТ НА RFI (по желанию)", "BOP: единый BPA по всей стране — химия + ежемесячные анализы, Legionella, онлайн-портал, свои люди на объектах. Как генподрядчик не тянем (лимит субподряда FAR 52.219-14). Имеет смысл короткий ответ: предложить выделить поставку химии в отдельный лот под NAICS 325998 — бесплатно, ставит нас в список рынка. Срок 03.12."),
 "N0010427QBC46": ("НЕТ", "Уголь для ВМФ: документы NOFORN (не для иностранных граждан) через сайт ядерной программы ВМФ, CMMC, лотовые испытания. Для владельца-гражданина РФ закрыто."),
 "N0010427QBC45": ("НЕТ", "Хромат калия для ядерных объектов ВМФ — то же: NOFORN, CMMC."),
 "W9123826X1CUS": ("ЛИД", "PFAS-изыскания USACE — инженерная услуга. Только как будущая субпоставка GAC генподрядчику."),
 "N4008527R0048": ("ЛИД", "Стройка станции очистки Bethpage NY — субпоставка GAC генподрядчику после выбора подрядчика."),
 "W9123827X1CHG": ("ЛИД", "Стройка очистных Calaveras CA — субпоставка химии после выбора подрядчика."),
}
rows = list(csv.DictReader(open("data/usa/sam_1711/sam_opportunities_2026-11-17.csv", encoding="utf-8-sig")))
wb = Workbook(); ws = wb.active; ws.title = "Разбор SAM 17.11"
hdr = ["Срок ответа", "Notice ID", "Название", "Ведомство", "Штат", "Тип", "Set-aside", "NAICS", "Решение", "Почему", "Ссылка"]
ws.append(hdr)
for c in ws[1]: c.fill = H; c.font = HF; c.alignment = W
for r in rows:
    nid = r["Notice ID"]
    v, why = VERDICT.get(nid, ("НЕТ", r["Релевантность / комментарий"]))
    ws.append([r["Срок ответа"], nid, r["Название"], r["Ведомство"], r["Штат"], r["Тип"], r["Set-aside"], r["NAICS"], v, why, r["Ссылка"]])
    ws.cell(ws.max_row, 9).fill = G if v == "ИДЁМ" else Y if v in ("ПРОВЕРИТЬ", "ОТВЕТ НА RFI (по желанию)", "ЛИД") else R
for i, w in enumerate([18, 20, 40, 30, 6, 18, 18, 9, 16, 90, 40], 1): ws.column_dimensions[chr(64 + i)].width = w
for row in ws.iter_rows(min_row=2):
    for c in row: c.alignment = W
ws.freeze_panes = "A2"

# --- 2. Rock Salt: экономика ----------------------------------------------------
QY = 210; YEARS = 3
HIST = [("M6700125P0005 KGOC Group (GA), 07.2025, 10 участников", 28718), ("M6700126P1068 Cultivated Investments (SC), 07.2026, 6 участников", 50400),
        ("M6700122P1011 Amerochem (NC), 2022–2025, 3 года, 2 участника", 133980)]
e = wb.create_sheet("Rock Salt — экономика")
e.append(["Rock Salt M6700127Q1014 — MCAS Cherry Point: 210 short tons/год, база + 2 опциона, пневмодоставка в подземный бак, мин. 25 т, 5 дней"]); e.cell(1, 1).font = Font(bold=True, size=12)
e.append([]); e.append(["Прошлые контракты (USAspending)", "Сумма, $", "$/т при 210 т/год"])
for c in e[3]: c.fill = H; c.font = HF
for n, s in HIST:
    t = QY * 3 if "3 года" in n else QY
    e.append([n, s, round(s / t)])
e.append(["Рынок: каменная соль с доставкой $80–130/sh t, солевая (solar) $100–160/sh t (USGS 2026 / отраслевые обзоры); +пневмовыгрузка ~$10–25/т"])
e.append([]); e.append(["Цена поставщика (с доставкой и пневмо), $/т", "Наша цена, $/т", "Выручка/год, $", "Себестоимость/год, $", "Накладные/год (банк, страховка, DBIDS, связь), $", "Прибыль/год, $", "Маржа, %", "Выручка за 3 года, $"])
for c in e[e.max_row]: c.fill = H; c.font = HF; c.alignment = W
OVH = 1500
for sup in (120, 140, 160):
    for our in (165, 185, 205):
        rev = our * QY; cost = sup * QY; pr = rev - cost - OVH
        e.append([sup, our, rev, cost, OVH, pr, round(pr / rev * 100, 1), rev * YEARS])
e.append([]); e.append(["Рекомендация: цена 175–195 $/т (между KGOC 137 и Cultivated 240). Окончательно — после КП поставщиков (запросы: Mid-Atlantic Salt, Rock Salt Today, Chemical Equipment Labs, Mid-American Salt)."])
e.column_dimensions["A"].width = 60
for col in "BCDEFGH": e.column_dimensions[col].width = 16
for row in e.iter_rows():
    for c in row: c.alignment = W

# --- 3. Чек-лист подачи ---------------------------------------------------------
k = wb.create_sheet("Rock Salt — чек-лист")
k.append(["№", "Что подать (Section L)", "Статус / кто"])
for c in k[1]: c.fill = H; c.font = HF
for x in [
 ("1", "Письмо-котировка: способность выполнить (поставщик с пневмовозом, срок 5 дней, SDS) — черновик QUOTE_M6700127Q1014.docx", "готово, вставить поставщика и цену"),
 ("2", "Цены по CLIN 0001 / 1001 / 1002: цена за тонну и сумма (210 т каждая), FOB Destination", "после КП поставщика"),
 ("3", "SF1449 стр.1: блоки 17a (название, адрес, CAGE 217F4, UEI, телефон) и 30a–c (подпись, ФИО, дата)", "Вы — подписать PDF"),
 ("4", "DFARS 252.204-7017 — представление по телеком-оборудованию (Huawei/ZTE и т.п.): «does not provide» (в SAM ответ 252.204-7016 «does not»)", "Вы — отметить в PDF"),
 ("5", "Past performance: 3 референса или прямое заявление «no relevant past performance» (оценивается нейтрально)", "в письме — заявление; при наличии реальных поставок — вписать"),
 ("6", "Тема письма: «RFQ M67001-27-Q-1014 — ISP Group LLC quote»; отправить на sierra.c.maye.civ@usmc.mil до 23.11 16:00 ET", "Вы — отправка"),
 ("7", "После победы: регистрация в WAWF/PIEE (оплата), DBIDS для водителей поставщика", "после award"),
 ("!", "Вопросы к заказчику закрылись 17.11 16:00 ET — спрашивать уже нельзя; поставлять ли каменную или выварочную (solar) соль — в ТЗ сказано «rock salt», котируем rock salt, solar — как альтернативу", "")]:
    k.append(list(x))
k.column_dimensions["A"].width = 4; k.column_dimensions["B"].width = 110; k.column_dimensions["C"].width = 40
for row in k.iter_rows():
    for c in row: c.alignment = W
wb.save(f"{OUT}/SAM_RAZBOR_1711.xlsx")

# --- 4. Черновик котировки ------------------------------------------------------
d = Document(); st = d.styles["Normal"]; st.font.name = "Calibri"; st.font.size = Pt(11)
P = d.add_paragraph
P("ISP Group LLC\n16395 Biscayne Blvd, North Miami Beach, FL 33160\nUEI: WA46PAMTYMQ7   CAGE: 217F4\nPhone: +1 929 830 0038   Email: info@ispgroupgc.com")
P("November [ ], 2026")
P("To: Sierra Maye, Contract Specialist\nMCIEAST-MCB Camp Lejeune / MCAS Cherry Point Satellite Contracting Office\nsierra.c.maye.civ@usmc.mil")
P("Subject: RFQ M67001-27-Q-1014 (M6700127Q1014) — Rock Salt, Central Heating Plant, MCAS Cherry Point — Quotation").runs[0].bold = True
P("ISP Group LLC, a small business registered and active in SAM (UEI WA46PAMTYMQ7), is pleased to submit this firm-fixed-price quotation for the supply, delivery and pneumatic off-loading of rock salt in accordance with the Purchase Description of RFQ M67001-27-Q-1014.")
d.add_heading("1. Pricing (FOB Destination, MCAS Cherry Point, NC 28533)", level=2)
t = d.add_table(rows=1, cols=5); t.style = "Table Grid"
for i, h in enumerate(["CLIN", "Description", "Qty", "Unit price, $/ton", "Amount, $"]): t.rows[0].cells[i].text = h
for cl, de in [("0001", "Rock Salt — base year"), ("1001", "Rock Salt — option year 1"), ("1002", "Rock Salt — option year 2")]:
    r = t.add_row().cells; r[0].text = cl; r[1].text = de; r[2].text = "210 tons"; r[3].text = "[ ]"; r[4].text = "[ ]"
r = t.add_row().cells; r[1].text = "TOTAL (base + options)"; r[2].text = "630 tons"; r[4].text = "[ ]"
P("Prices include product, freight, fuel surcharge, pneumatic off-loading and all delivery charges. Quotation valid for 60 days after the closing date.")
d.add_heading("2. Technical ability", level=2)
for x in ["Product: [rock salt / water-conditioning grade, origin: (mine/terminal), USA], sodium chloride ≥ [ ]% NaCl; Certificate of Analysis and Safety Data Sheet attached; SDS will be provided at the delivery site (PD 2.3).",
          "Delivery: full truckloads of 25 tons minimum, delivered within 5 calendar days of each call from the Government authorized caller, between 7:30 AM and 3:00 PM ET (PD 2.1, 2.4).",
          "Off-loading: pneumatic bulk tankers with on-board blowers, capable of blowing salt through the 4-inch fill line into the underground brine tank at Building 152 (PD 2.2). Drivers will obtain DBIDS credentials prior to the first delivery.",
          "Supply chain: product sourced from [supplier name], a bulk salt distributor operating pneumatic delivery equipment in North Carolina, under a supply agreement with ISP Group LLC; ISP Group LLC is the prime contractor responsible for ordering, scheduling, quality, invoicing via WAWF and all communication with the Government.",
          "Safety and compliance: the contractor will provide all required safety equipment, comply with Federal, State and local regulations and installation traffic rules, and is responsible for any damage to Government property (PD 4.0).",
          "Buy American: the end product is a domestic end product (U.S.-mined salt)."]:
    d.add_paragraph(x, style="List Bullet")
d.add_heading("3. Past performance", level=2)
P("ISP Group LLC affirmatively states that it possesses no relevant past performance history with Federal contracts of similar technical complexity. [If applicable, list commercial supply references here: customer, contract no., POC, description, value, period.] Our distribution partner [supplier] has an established record of pneumatic bulk salt deliveries to municipal and industrial water plants in the region.")
d.add_heading("4. Representations", level=2)
P("Completed SF1449 (blocks 17 and 30) and DFARS 252.204-7017 representation are attached. Annual representations and certifications are current in SAM.")
P("Sincerely,\n\n______________________\nAleksandr Orlov, Managing Member\nISP Group LLC")
d.save(f"{RS}/QUOTE_M6700127Q1014.docx")
print("ok", len(rows))
