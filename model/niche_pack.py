"""Ниша «катализаторы / адсорбенты / реагенты» (как у СПб НХК и БПО Евразия-Снаб): реестр всех найденных конкурсов + производители под каждый + запросы на дилерство.
Вход: data/companies/{eurasia_out,spbnhk_out,dealers_out}.json. Выход: ../NISHA_KONKURSY_DILERSTVA.xlsx, ../NISHA_KONKURSY_DILERSTVA.pdf, ../ZAPROSY_DILERSTVO.docx"""
import json, html, re, openpyxl, docx
from docx.shared import Pt
from openpyxl.styles import Font, PatternFill, Alignment
from pdfutil import html_to_pdf
from tidy import tidy

D = "data/companies/"
EU = json.load(open(D + "eurasia_out.json")); NK = json.load(open(D + "spbnhk_out.json")); DL = json.load(open(D + "dealers_out.json"))
W = DL["where_to_get"]
# сегменты → индексы производителей в dealers_out.where_to_get (+ кто ещё может, если в списке нет)
SEG = [
    ("Продажа заказчиком (реализация отработанного) — не поставка", r"реализаци|отработанн", []),
    ("Услуги: перезагрузка катализаторов", r"перезагрузк", [0, 1]),
    ("Оборудование: реакторы, запчасти компрессоров, массообмен", r"реактор|компрессор|запчаст|запасн|ЗИП|насадк|тарелк|массообмен|когенерац", [27, 28]),
    ("Катализаторы НПЗ (гидроочистка, риформинг, водород, защитный слой, сероочистка)", r"катализ.*(гидро|риформ|водород|защитн|парекс|изомер|крекинг|л-35|сернист)|риформинг", [0, 1, 5, 7, 13, 12]),
    ("Катализаторы азот/метанол/нефтехимия", r"катализ|DQC|ZnO|цинкоокс", [6, 13, 5, 7, 1]),
    ("Адсорбенты, цеолиты, силикагель, оксид алюминия, керамические шары, носители", r"адсорб|цеолит|молекуляр|сито|сита|оксид алюмин|силикагел|керамическ|осушит|Tonsil|Linde|носител|сорбент", [1, 0, 14, 15, 7, 13]),
    ("Активированный уголь, иониты", r"угл|уголь|ионит|ионообмен|смол", [8]),
    ("Реагенты нефтедобычи (деэмульгаторы, ПТП, поглотители H2S)", r"деэмульг|противотурбулент|ПТП|сероводород|ЦППСН|бактерицид", [19, 20]),
    ("Реагенты водоподготовки энергетики (ингибиторы, гидразин, поглотители кислорода, ВХО)", r"ингибитор|гидразин|кислород|Инкорамин|Оптион|реагент|водно-хим|ВХО|накип|аммиак", [21, 22, 19]),
    ("Гликоли, растворители, присадки, прочая химия", r"гликол|раствор|присадк|метанол", [21, 19]),
]
def seg(text):
    for name, rx, idx in SEG:
        if re.search(rx, text, re.I): return name, idx
    return "Прочее", []
rows = {}
for src, L in (("БПО Евразия-Снаб", EU["current_tenders"]), ("СПб НХК", NK["current_tenders"])):
    for t in L:
        k = str(t.get("number")).strip()
        if k in rows: rows[k]["found_for"] += ", " + src; continue
        s, idx = seg(str(t.get("subject")))
        nmc = t.get("nmc") if isinstance(t.get("nmc"), (int, float)) else None
        rows[k] = dict(t, number=k, nmc=nmc, found_for=src, segment=s, makers="; ".join(W[i]["manufacturer"].split(" (")[0] for i in idx),
                       makers_contacts="\n".join(f"{W[i]['manufacturer'].split(' (')[0]}: {(W[i].get('contacts') or {}).get('email') or ''} {(W[i].get('contacts') or {}).get('phone') or ''}".strip() for i in idx),
                       eis="ЕИС" if re.fullmatch(r"\d{11}", k) else "площадка заказчика / агрегатор")
R = sorted(rows.values(), key=lambda r: str(r.get("deadline")))

# ─── Excel ───
wb = openpyxl.Workbook(); ws = wb.active; ws.title = f"Конкурсы ({len(R)})"
cols = [("№ закупки", "number"), ("Где", "eis"), ("Заказчик", "customer"), ("Предмет", "subject"), ("НМЦ, ₽", "nmc"), ("Подача до", "deadline"), ("Сегмент", "segment"),
        ("Под чей профиль найден", "found_for"), ("Соответствие / комментарий", "fit"), ("Производители — запросить письмо/цену", "makers"), ("Контакты производителей", "makers_contacts"), ("Ссылка", "url")]
ws.append([c for c, _ in cols])
for r in R:
    ws.append([r.get(k) if k == "nmc" else str(r.get(k) or "") for _, k in cols])
    X = ws.max_row; ws.cell(X, 5).number_format = "#,##0"
    u = str(r.get("url") or "").split(" ")[0]
    if u.startswith("http"): ws.cell(X, 12).hyperlink = u
for c in ws[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
for i, w in enumerate([15, 12, 30, 50, 14, 16, 30, 16, 40, 40, 50, 50], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "C2"; ws.auto_filter.ref = f"A1:L{ws.max_row}"
s2 = wb.create_sheet("Сводка по сегментам"); s2.append(["Сегмент", "Конкурсов", "Σ НМЦ (где известна), ₽", "Ближайший срок", "Производители"])
for name, _, idx in SEG + [("Прочее", "", [])]:
    L = [r for r in R if r["segment"] == name]
    if L: s2.append([name, len(L), sum(r["nmc"] or 0 for r in L), min(str(r.get("deadline")) for r in L), "; ".join(W[i]["manufacturer"].split(" (")[0] for i in idx)])
for row in s2.iter_rows(min_row=2): row[2].number_format = "#,##0"
s3 = wb.create_sheet("Производители (дилерство)"); c3 = ["segment", "manufacturer", "country", "inn", "dealer_program", "site", "phone", "email", "verified", "fits", "why"]
s3.append(["Сегмент", "Производитель", "Страна", "ИНН", "Дилерская программа", "Сайт", "Телефон", "E-mail", "Контакты подтверждены", "Кому подходит", "Почему"])
for w in W:
    c = w.get("contacts") or {}
    s3.append([str(c.get(k) or "") if k in ("site", "phone", "email") else str(w.get(k) or "") for k in c3])
s4 = wb.create_sheet("Как СПб НХК и БПО работают"); s4.append(["Факт", "Источник"])
for a, b in [("Обе — перепродавцы: сами ничего не производят; оборот СПб НХК 1,4–2,0 млрд/год, БПО — 0,7 млрд (2025)", "досье, saby.ru"),
             ("В ЕИС почти не участвуют: продажи идут через закрытые площадки нефтяных компаний (Роснефть, Газпром нефть, ЛУКОЙЛ, СИБУР) и прямые договоры", "досье"),
             ("Дилерства публично не подтверждаются; у СПб НХК на сайте устаревший статус «Зульцер Хемтех» (теперь «Сплитекс»)", "проверка дилерств"),
             ("Ценность для заказчика: подбор аналога (импортозамещение ушедших JM/UOP/Evonik), китайские поставки, шеф-монтаж загрузки катализатора", "сайты, досье")]: s4.append([a, b])
for s in (s2, s3, s4):
    for c in s[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
    for i in range(1, 12): s.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 30
tidy(wb, row_h=36); wb.save("../NISHA_KONKURSY_DILERSTVA.xlsx")

# ─── PDF (короткий) ───
e = lambda s: html.escape(str(s if s is not None else "—"))
fm = lambda v: f"{v:,.0f}".replace(",", " ") if isinstance(v, (int, float)) else "—"
css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:7.8px;color:#111;margin:0}h1{font-size:13px;margin:0 0 3px}h2{font-size:10px;margin:5px 0 2px;border-bottom:2px solid #305496}
table{border-collapse:collapse;width:100%}th,td{border:1px solid #bbb;padding:1px 3px;vertical-align:top}th{background:#305496;color:#fff;text-align:left}td.n{text-align:right;white-space:nowrap}.small{font-size:6.8px;color:#444}"""
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body><h1>Ниша «катализаторы, адсорбенты, реагенты»: все найденные конкурсы ({len(R)}) и у кого брать товар</h1>",
     "<div class='small'>Конкурсы собраны при разборе СПб НХК и БПО Евразия-Снаб (подача с 06.11.2026). НМЦ по закупкам с агрегаторов часто скрыта или указана за единицу. Контакты производителей и тексты запросов на дилерство — в Excel и Word.</div><h2>По сегментам</h2><table><tr><th>Сегмент</th><th>Конкурсов</th><th>Σ НМЦ (известная)</th><th>Производители</th></tr>"]
for row in s2.iter_rows(min_row=2, values_only=True): h.append(f"<tr><td>{e(row[0])}</td><td class=n>{row[1]}</td><td class=n>{fm(row[2])}</td><td class=small>{e(row[4])}</td></tr>")
h.append("</table><h2>Конкурсы по срокам</h2><table><tr><th>Подача до</th><th>№</th><th>Заказчик</th><th>Предмет</th><th>НМЦ</th><th>Сегмент</th></tr>")
for r in R: h.append(f"<tr><td>{e(str(r.get('deadline'))[:16])}</td><td class=small>{e(r['number'])}</td><td class=small>{e(r.get('customer'))[:55]}</td><td>{e(r.get('subject'))[:110]}</td><td class=n>{fm(r['nmc'])}</td><td class=small>{e(r['segment'])[:45]}</td></tr>")
h.append("</table></body></html>")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/nisha.html"
open(hp, "w").write("".join(h)); html_to_pdf(hp, "../NISHA_KONKURSY_DILERSTVA.pdf")

# ─── Word: запросы на дилерство ───
LET = [0, 1, 6, 7, 8, 5, 19, 21, 27, 28, 13, 14, 15, 18]   # производители, которым имеет смысл писать
doc = docx.Document(); st = doc.styles["Normal"]; st.font.name = "Times New Roman"; st.font.size = Pt(11)
doc.add_heading("Запросы на дилерство / авторизацию — шаблоны", 1)
doc.add_paragraph("Вставьте реквизиты своей компании вместо [ ]. Отправлять в отдел продаж/сбыта; после ответа — звонок и встреча. Китайским и прочим иностранным производителям — английский текст.")
for i in LET:
    w = W[i]; c = w.get("contacts") or {}; name = w["manufacturer"].split(" (")[0]
    foreign = not str(w.get("country", "")).startswith("РФ")
    doc.add_page_break(); doc.add_heading(name, 2)
    doc.add_paragraph(f"Кому: {c.get('email') or 'e-mail уточнить'}; тел. {c.get('phone') or '—'}; сайт {c.get('site') or '—'}. Сегмент: {w.get('segment')}.")
    if not foreign:
        doc.add_paragraph(f"Тема: Запрос на заключение дилерского (дистрибьюторского) договора — {w.get('segment').split('(')[0].strip()}")
        doc.add_paragraph(f"Уважаемые коллеги!\n\nООО «[наименование]» (ИНН [ ], г. [ ]) — поставщик материально-технических ресурсов для предприятий энергетики, нефтегазовой и химической отрасли; участвуем в закупках по 223-ФЗ и 44-ФЗ "
                          f"(ПАО «Т Плюс», ПАО «ТГК-1», ПАО «Россети», ПАО «Форвард Энерго», нефтедобывающие и нефтеперерабатывающие предприятия) и на корпоративных площадках.\n\n"
                          f"Просим рассмотреть возможность заключения с нами дилерского договора на продукцию {name} — {w.get('segment')}. В ближайшие недели планируем участие в закупках по этой номенклатуре "
                          f"(перечень с номерами закупок и сроками прилагаем) и готовы предлагать именно вашу продукцию.\n\nПросим сообщить:\n"
                          "1) условия дилерской программы (объёмы, скидки от прайса, порядок оплаты/отсрочки);\n2) актуальный прайс-лист и сроки изготовления/отгрузки;\n"
                          "3) возможность выдачи авторизационного письма производителя под конкретную закупку (номер закупки, заказчик, наименование продукции) и срок его оформления;\n"
                          "4) наличие записей в реестре российской промышленной продукции (ПП 719) и заключений аттестации ПАО «Россети» (при наличии);\n"
                          "5) комплект документов для заявок: ТУ, паспорта качества, сертификаты/декларации ЕАЭС, паспорта безопасности.\n\n"
                          "Со своей стороны предоставим карточку предприятия, сведения о закупках и прогноз объёмов на 2027 г.\n\nС уважением,\n[должность, ФИО]\n[телефон, e-mail]")
    else:
        doc.add_paragraph(f"Subject: Request for distributorship / authorization in Russia — {name}")
        doc.add_paragraph(f"Dear Sirs,\n\n[Company name] LLC (Russia, St. Petersburg, TIN [ ]) supplies materials to Russian power generation, oil & gas and chemical companies via public tenders (Law 223-FZ/44-FZ) and corporate procurement platforms.\n\n"
                          f"We would like to become your distributor in Russia for {w.get('segment')}. We are currently preparing bids where your products fit (list of tenders with numbers and deadlines attached).\n\n"
                          "Kindly advise:\n1) terms of your distributor program (volumes, discounts, payment terms in CNY);\n2) current price list (FCA/CIF/DAP Russia), production and delivery lead time;\n"
                          "3) possibility to issue a Manufacturer Authorization Letter for a specific tender (tender number, end user, product), apostilled, with Russian translation, and the lead time to issue it;\n"
                          "4) available documents: TDS, MSDS, CoA, quality certificates, references in Russia/CIS;\n5) contact person for Russia/CIS.\n\nBest regards,\n[Name, position]\n[phone, e-mail, WeChat]")
    doc.add_paragraph(f"Почему этот производитель: {w.get('why')}")
doc.save("../ZAPROSY_DILERSTVO.docx")
print(len(R), {s: sum(1 for r in R if r['segment'] == s) for s in set(r['segment'] for r in R)})
