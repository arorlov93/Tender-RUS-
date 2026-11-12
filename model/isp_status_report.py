"""Сводка ISP Group на 12.11.2026: что можем поставлять (поставщики, дилерство) и переписка с заказчиками/тендерами.
Источник: почта info@ispgroupgc.com (письма 08–11.11.2026), журналы рассылок data/companies/*sent*.json. → ../ISP_STATUS_PRODUCTS_TENDERS_1211.pdf + .xlsx"""
import html, os, subprocess
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from isp_letterhead import SCRATCH

# Что можем поставлять
PROD = [
 ("Фосфонаты: HEDP 60/90%, ATMP 50/95%, PBTCA 50%, DTPMPA 50%; антискалант RO KR-0100; цинк-фосфонатные ингибиторы", "Zaozhuang Kairui Water Treatment (Китай), Винсент Ян",
  "ЭКСКЛЮЗИВ Туркменистан (LOA получен, обновлённый LOA ждёт подписи ГД). Казахстан: неэксклюзивно по тем же ценам. Договор v2 (200 т/год) Kairui принял, подпишет ГД после возвращения (на след. неделе)",
  "FOB Qingdao 11.11: ATMP 50% $1 195/т; HEDP 60% $1 315/т; PBTCA 50% $1 345/т. CIF Nhava Sheva: $1 370 / $1 490 / $1 520", "100% T/T против копии BL; FCL 25 т (IBC 1 250 кг); цена 5 дней"),
 ("Фосфонаты (альтернатива): HEDP 60%, PBTCA 50%, ATMP 50%", "Zaozhuang Jialu Chemical (Китай), Rachel Li",
  "Только КП, дилерство не обсуждали", "FOB Qingdao: HEDP $1 410; PBTCA $1 500; ATMP $1 300/т (CIF NS +$160)", "50% аванс, 50% против копии BL; 10 дней; цена до 15.11"),
 ("Ионообменные смолы: катионит 001x7 (≈КУ-2-8) H/Na; анионит 201x7 (≈АВ-17-8); D301", "Jiangsu Suqing Water Treatment (Китай), Haiyan Cha",
  "Готовы рассмотреть дистрибьюторский договор СНГ/ЦА — мы предложили начать 10.11, ответа пока нет", "FOB Shanghai: 001x7 H $840/м³; Na $860/м³; 201x7 Cl $2 090/м³; D301 $2 100/м³", "30% аванс, 70% против копии BL; MOQ 3 м³/поз.; 20–30 дней"),
 ("Полиоксихлорид алюминия (PAC) порошок: питьевой (NSF) и технический", "Henan Aierfuke Chemicals (Китай), Kevin Xu",
  "Только КП + документы (NSF, ISO 9001/14001/45001, COA, MSDS). Уже отгружают в СНГ/ЦА. Дилерство не обсуждали", "FOB Qingdao: технический $295/т; питьевой $355/т (без паллет)", "100% T/T аванс; MOQ 1×20' (22 т); 7–10 дней"),
 ("Деэмульгаторы (нефте- и водорастворимые), их полупродукты; ингибиторы коррозии", "Rimpro India, Govind Patel",
  "Согласились выдать письмо-авторизацию для тендеров Туркменистана; черновик LOA отправлен 10.11. Цены после праздников (до 17.11)", "Цен пока нет", "Условия оплаты получены (в переписке)"),
 ("Молекулярные сита, активный оксид алюминия", "Jalon (Китай), Siyee",
  "КП и документы пришлют на этой неделе (ответ 11.11). Схема оплаты из США на счёт компании подтверждена с нашей стороны", "Цен пока нет", "—"),
 ("Фосфонаты: прочие производители", "Xintai (автоответ, 24 ч), Taihe, Jiangsu Yao's, Koye, Qingshuiyuan, IRO, Guiquan, Yuanlian",
  "Запросы цен отправлены 11.11, ответов пока нет", "—", "—"),
]
# Заказчики / тендеры
BUY = [
 ("Chembond (Индия)", "Запрос: HEDP 60% ~600 т/год, PBTC 50% ~400 т/год, CIF Nhava Sheva", "11.11: подтвердили, готовим КП CIF Nhava Sheva",
  "Китайские цены выше индийского рынка (≈$1 000/т CIF HEDP): Kairui $1 490, Jialu $1 570. Ждём других заводов; по HEDP вероятно не конкурентны, предложим PBTCA"),
 ("Petronas (Малайзия), регистрация поставщика PC1", "Регистратор ответил на 2 наших вопроса: «Confirmed / Acceptable» (11.11)", "Анкеты и KYC готовы",
  "Нужны от вас: подписанная Compliance Policy, Articles of Organization, выписка Sunbiz, письмо и выписки Bank of America"),
 ("Oyu Tolgoi (Монголия)", "Попросили зарегистрироваться в OASIS", "11.11: сообщили, что зарегистрировались и подали профиль (Process Chemicals)", "Ждём валидации профиля в OASIS"),
 ("АО «Алмалыкский ГМК» (Узбекистан)", "Попросили официальное письмо на имя гл. инженера", "10.11: отправлено письмо ISP-1011-02 с подписью (смолы, фосфонаты)", "Ждём ответа"),
 ("СП «Uz-Kor Gas Chemical» (Узбекистан), тендер T-47/2026", "Тендерную инструкцию дают по официальному запросу; подача до 02.12.2026", "11.11: отправлен запрос ISP-1111-02 с подписью", "Ждём документацию → готовим предложение"),
 ("ТОО «Karabatan Utility Solutions» (Казахстан)", "Попросили официальное письмо на имя Председателя Правления", "11.11: отправлено письмо ISP-1111-01 с подписью", "Ждём потребность и порядок закупок"),
 ("Шуртанский ГХК (Узбекистан)", "—", "11.11: письмо + тендерный ящик (как попасть в приглашённые)", "Ждём"),
 ("Туркменистан: Turkmengaz T/GAZ-038 (SY-210, ~14 т), др.", "—", "Письма заказчикам (6) и запросы LOA заводам (3), 10.11", "Тендер T/GAZ-038 решили пропустить; LOA Kairui и Rimpro в работе"),
]
STATS = [("10.11 мир (смолы, фосфонаты)", 25), ("10.11 волна 2 (PAC и др.)", 55), ("10.11 Туркменистан", 6), ("11.11 волна 3 (КЗ/УЗ 19 + формуляторы 24)", 43)]

e = html.escape
css = """@page{size:A4 landscape;margin:8mm}body{font-family:'DejaVu Sans',Arial;font-size:9.5px;color:#111}h1{font-size:13px;margin:0;color:#0B2545}h2{font-size:10px;margin:8px 0 3px;color:#0B2545}
table{border-collapse:collapse;width:100%}th{background:#0B2545;color:#fff;padding:3px;text-align:left;vertical-align:top}td{border-bottom:1px solid #ddd;padding:2px;vertical-align:top}.g{color:#555}b.ok{color:#1B7F3B}"""
h = [f"<!DOCTYPE html><html><head><meta charset=utf-8><style>{css}</style></head><body><h1>ISP Group LLC — продукты, дилерство и переписка с заказчиками · 12.11.2026</h1>",
     "<div class=g>Источник: почта info@ispgroupgc.com (08–11.11.2026). Цены поставщиков — для нас, в письмах заказчикам не раскрываются.</div>",
     "<h2>1. Что можем поставлять</h2><table><tr><th style='width:20%'>Продукция</th><th style='width:14%'>Производитель</th><th style='width:27%'>Дилерство / авторизация</th><th style='width:22%'>Цена (последняя)</th><th>Условия</th></tr>"]
for r in PROD: h.append("<tr>" + "".join(f"<td>{e(x)}</td>" for x in r) + "</tr>")
h.append("</table><h2>2. Заказчики и тендеры: кто ответил, что мы ответили</h2><table><tr><th style='width:20%'>Заказчик / тендер</th><th style='width:24%'>Что ответили</th><th style='width:24%'>Что мы ответили</th><th>Статус / следующий шаг</th></tr>")
for r in BUY: h.append("<tr>" + "".join(f"<td>{e(x)}</td>" for x in r) + "</tr>")
h.append("</table><h2>3. Рассылки заказчикам</h2><div>" + " · ".join(f"{e(a)}: <b>{b}</b>" for a, b in STATS) + f" · <b>Итого {sum(b for _, b in STATS)} писем</b>. Возвраты 11.11: 2 из 43 (Horizon — блок по IP, Granimex — нет адреса). Завтра: 5 писем формуляторам.</div>")
h.append("<h2>4. Что нужно от вас</h2><div>1) Документы для Petronas (см. выше). 2) Решение по Chembond: предлагать только PBTCA или ждать цен других заводов. 3) Подпись договора Kairui — после подписи их ГД.</div></body></html>")
hp = f"{SCRATCH}/status.html"; open(hp, "w").write("".join(h))
subprocess.run(["/opt/pw-browsers/chromium", "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer", f"--print-to-pdf={os.path.abspath('../ISP_STATUS_PRODUCTS_TENDERS_1211.pdf')}", "file://" + hp], check=True, capture_output=True, timeout=180)

wb = Workbook(); H = PatternFill("solid", fgColor="0B2545"); HF = Font(bold=True, color="FFFFFF"); wrap = Alignment(wrap_text=True, vertical="top")
for title, head, rows, widths in [("Продукты", ["Продукция", "Производитель", "Дилерство / авторизация", "Цена", "Условия"], PROD, [40, 30, 55, 45, 35]),
                                  ("Заказчики", ["Заказчик / тендер", "Что ответили", "Что мы ответили", "Статус / шаг"], BUY, [35, 45, 45, 50])]:
    ws = wb.active if title == "Продукты" else wb.create_sheet(title); ws.title = title; ws.append(head)
    for c in ws[1]: c.fill = H; c.font = HF
    for r in rows: ws.append(list(r))
    for row in ws.iter_rows(min_row=2):
        for c in row: c.alignment = wrap
    for i, w in enumerate(widths): ws.column_dimensions["ABCDE"[i]].width = w
ws = wb.create_sheet("Рассылки"); ws.append(["Волна", "Писем"]); [ws.append(list(r)) for r in STATS]
wb.save("../ISP_STATUS_PRODUCTS_TENDERS_1211.xlsx"); print("ok")
