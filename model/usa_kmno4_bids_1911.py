"""KMnO4 bids USA (19.11.2026): Magnesia CIF Houston $3,451/t (25 kg pails), $3,653/t (150 kg drums) → ../ZAYAVKI_USA/KMNO4/KMNO4_BIDS_1911.xlsx"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
LB = 2.20462
def landed(cif_t, fr_lb):
    cif = cif_t / 1000; fob = cif - 7500 / 18000
    kg = cif + fob * 0.155 + cif * 0.0047 + 0.075 + 0.06   # duty 5.5%+10% India, MPF/HMF, port/broker, warehouse
    return kg / LB + fr_lb                                  # $/lb delivered
B = [  # entity, state, due, submission, pack, qty_lb, prior prices, freight $/lb from Houston (estimate), bid $/lb, notes
 ("Hardin County Water District No.1", "KY", "25.11.2026 14:00 ET", "e-mail cgohman@hcwd.com (тема «2027 WATER/WASTEWATER TREATMENT CHEMICALS BID»)", "55 lb drums (=25 kg pails)", None, "2026: 2,28–5,00 $/lb; 2025 Brenntag 2,38", 0.20, 2.45, "Цена за барабан и за lb, доставка с lift gate на 5 станций, цена твёрдая на 2027. Бланк: Chem-Bid_Form_2027.pdf, подаём только строку KMnO4."),
 ("City of Oregon OH – Water Treatment", "OH", "01.12.2026 11:00 ET", "бумага, 5330 Seaman Rd, Oregon OH; bid bond/чек $500", "по спецификации", 24000, "Lorain OH 2026: 2,19–2,37 $/lb", 0.20, 2.40, "12 т KMnO4. Нужен залог $500 и бумажный пакет (курьер)."),
 ("Saratoga County NY – Central Services", "NY", "02.12.2026 11:00 ET", "Empire State Purchasing Group (empirestatebidsystem.com) или лично", "уточнить", None, "нет данных", 0.25, 3.20, "Тендер только на KMnO4, 1 год + опция. Скачать пакет с empirestatebidsystem.com (регистрация)."),
 ("City of Alexander City AL – Utilities", "AL", "14.12.2026 14:30 CT", "только почта/курьер: City Clerk, 281 James D Nabors Dr, Alexander City AL 35010", "по пакету", None, "нет данных", 0.15, 3.20, "Нужно: страховка GL ≥ $1 млн (сертификат с заявкой), E-Verify affidavit, торговые референции."),
 ("City of Lake Charles LA – Water", "LA", "15.12.2026 14:45 CT", "BidNet Direct (электронно) или бумага", "330,75 lb drums (=150 kg)", 9600, "нет данных", 0.08, 3.30, "Страховки: GL $1 млн, auto $500 тыс., WC $100 тыс.; местное предпочтение. Близко к Хьюстону."),
 ("City of Craig CO", "CO", "15.12.2026 10:00 MT", "e-mail cityclerk@cityofcraig.org / BidNet / почта", "55 lb plastic buckets", 10000, "нет данных", 0.25, 3.40, "10 000 lb/год, партии по 2 000 lb, доставка в Craig включена."),
]
H = PatternFill("solid", fgColor="0B2545"); HF = Font(bold=True, color="FFFFFF"); W = Alignment(wrap_text=True, vertical="top")
wb = Workbook(); ws = wb.active; ws.title = "Закупки KMnO4"
ws.append(["Заказчик", "Штат", "Подача", "Как подавать", "Фасовка", "Объём 2027, lb", "Прошлые цены", "Себестоимость с доставкой, $/lb", "Предлагаемая цена, $/lb", "Маржа, %", "Выручка/год, $", "Что нужно"])
for c in ws[1]: c.fill = H; c.font = HF; c.alignment = W
for e, st, due, sub, pack, q, prior, fr, bid, note in B:
    cif = 3653 if "330" in pack else 3451
    cost = landed(cif, fr)
    ws.append([e, st, due, sub, pack, q or "не указан", prior, round(cost, 2), bid, round((bid - cost) / bid * 100, 1), round(q * bid) if q else "—", note])
for i, w in enumerate([30, 6, 18, 40, 20, 12, 28, 14, 12, 9, 12, 60], 1): ws.column_dimensions[chr(64 + i)].width = w
for r in ws.iter_rows(min_row=2):
    for c in r: c.alignment = W
s = wb.create_sheet("Что сделать")
for x in ["1. Magnesia: продлить цену на 60 дней, NSF-листинг завода, COA по AWWA B603, смешанная загрузка (вёдра 25 кг + бочки 150 кг) — письмо в черновиках.",
          "2. Сроки: экспортная лицензия Индии 3–4 недели + 10–15 дней + ~35–40 дней море → первый контейнер в Хьюстоне ≈ середина февраля 2027. Контракты с 01.01.2027 — в заявках указать «first delivery from 15 Feb 2027» или договориться о стартовом запасе.",
          "3. США: DEA — уведомление об импорте (форма DEA-486, KMnO4 — List II), PHMSA регистрация ($275, окислитель кл. 5.1), склад 3PL с допуском hazmat в Хьюстоне, LTL-перевозчик hazmat.",
          "4. Страховка: GL $1/2 млн, auto $500 тыс.+, WC — нужна для Alexander City и Lake Charles; сертификат COI у брокера.",
          "5. Документы в заявку: бланк заказчика, W-9 ISP Group, NSF 60 листинг, SDS, TDS, AWWA B603 COA, референции.",
          "6. Hardin KY — 25.11 по e-mail: заполнить строку KMnO4 (цена за барабан 55 lb и за lb). Остальные строки оставить пустыми / «No bid»."]:
    s.append([x])
s.column_dimensions["A"].width = 150
for r in s.iter_rows():
    for c in r: c.alignment = W
wb.save("../ZAYAVKI_USA/KMNO4/KMNO4_BIDS_1911.xlsx")
for e, st, due, sub, pack, q, prior, fr, bid, note in B:
    cif = 3653 if "330" in pack else 3451; c = landed(cif, fr); print(st, due, round(c, 2), bid, round((bid - c) / bid * 100, 1))
