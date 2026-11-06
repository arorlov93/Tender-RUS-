"""13 приоритетных лотов (A/B на 06.11): таблица конкурсов + спецификации запроса КП по группам товаров. → ../KONKURSY_KP_0611.xlsx + ../ZAPROSY_KP_0611.docx"""
import json, re, openpyxl, docx
from docx.shared import Pt, Cm
from docx.enum.section import WD_ORIENT
from openpyxl.styles import Font, PatternFill, Alignment
from tidy import tidy
S = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/live/"
L = json.load(open(S + "items13.json"))
MG = {t["number"]: t for t in json.load(open("data/companies/niche_margin.json"))["tenders"]}
PE = json.load(open("data/companies/partners_energy.json"))
GROUPS = [  # группа, регэксп, поставщики для КП (название — контакт)
 ("Иониты (ионообменные смолы)", r"катионит|анионит|АН-31|КУ-2|ИОС", [("ООО «ТД Токем» (Кемерово)", "td.tokem@mail.ru, +7 (3842) 32-51-29"), ("ООО «АкваХимСервис» (Москва, дилер Токем)", "+7 (495) 545-42-74"), ("ООО «Интерхим Ком» (Sunresin, Москва)", "+7 (499) 130-22-33"), ("ООО «Фрагмит» (СПб)", "+7 (812) 242-78-42")]),
 ("Активированные угли", r"уголь|угля|Silcarbon", [("АО «Сорбент» (Пермь)", "carbon.sorbent@zelinskygroup.com, +7 (342) 258-61-74"), ("ЗАО «ЭХЗ» (Великий Новгород)", "+7 (8162) 67-53-13"), ("ООО «Уралхимсорб» (Пермь)", "контакт — через сайт")]),
 ("Ингибиторы коррозии и солеотложений", r"ингибитор|Позиция номер", [("ООО «Экоэнерго» (Оптион, Эктоскейл)", "op@ekoenergo.ru"), ("ООО «ТехноХимРеагентБел» (Гродно, ЕАЭС)", "+375 152 554054"), ("ООО «Компания Налко» (CINALAB)", "info@cinalab.ru, +7 495 980-70-60"), ("ООО НПФ «Эколарис» (Уфа)", "через сайт")]),
 ("ОЭДФ (фосфонаты)", r"ОЭДФ|HEDP", [("ООО «Экоэнерго»", "op@ekoenergo.ru"), ("Shandong Taihe Water Treatment (Китай)", "через сайт / Alibaba"), ("Henan Qingshuiyuan (Китай)", "через сайт")]),
 ("Аммиак водный технический", r"аммиак", [("ПАО «Акрон» (В. Новгород)", "через сайт acron.ru"), ("АО «НАК «Азот» (Новомосковск)", "через сайт"), ("АО «Тольяттиазот»", "через сайт"), ("ПАО «КуйбышевАзот»", "через сайт")]),
 ("Коагулянты (ПОХА)", r"полиоксихлорид|коагулянт", [("АО «Аурат» (Москва)", "office@aurat.ru, (499) 553-91-67"), ("ООО «Химпром» (Кемерово)", "tdx@extream.ru"), ("АО «Реагенты Водоканала» (Екатеринбург)", "через сайт"), ("ООО «Аква-Реагент»", "через сайт")]),
 ("Флокулянты и реагенты водоподготовки", r"флокулянт|реагентов для водоподготовки", [("ООО «Экоиндустрия» (ПАА, ПП 719)", "через сайт"), ("АО «Аурат»", "office@aurat.ru"), ("SNF (China) Flocculant", "через сайт snf")]),
 ("Пропиленгликоль / теплоносители", r"гликол", [("ООО «Химпром» (Кемерово), теплоноситель «Комфорт»", "tdx@extream.ru"), ("АО «СИБУР-Нефтехим» / ПАО «Нижнекамскнефтехим»", "через сайт"), ("Дистрибьюторы теплоносителей (Термагент, Dixis)", "уточнить")]),
]
def grp(t):
    for g, rx, sup in GROUPS:
        if re.search(rx, t, re.I): return g
    return "Прочее"
SUP = {g: sup for g, _, sup in GROUPS}
rows = []
for n, v in L.items():
    for it in v["items"]:
        name = it["info"] if it["info"] and not it["info"].startswith("Позиция номер") else v["title"]
        qty = float(it["qty"]) if it.get("qty") else None
        nmc = float(v["nmc"] or it.get("sum") or 0)
        term = "2027 г. (по заявкам заказчика)" if re.search(r"2027", v["title"] or "") else "в документации закупки — запросить ТЗ"
        g = grp((v["title"] or "") + " " + (name or ""))
        m = MG.get(n, {})
        rows.append(dict(n=n, cust=v["cust"], title=v["title"], name=name, qty=qty, unit=it.get("unit"), nmc=nmc, per=(nmc / qty if qty else None), close=str(v["close"])[:16], etp=v["etp"], method=v["method"],
                         msp="да" if "МСП" in str(v["method"]) or it.get("msp") == "true" else "", nat=("ограничение" if it.get("constr") == "true" else "") + (" преимущество" if it.get("sup") == "true" else ""),
                         place=it.get("place"), okpd=f"{it.get('okpd')} {it.get('okpd_name') or ''}", term=term, grp=g, marg=m.get("margin_m10"), docs=", ".join(x for x in v["docs"] if x)))
rows.sort(key=lambda r: r["close"])
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Конкурсы (13)"
H = ["Подача до", "№ ЕИС", "Заказчик", "Предмет", "Кол-во", "Ед.", "НМЦ, ₽", "НМЦ за ед., ₽", "Срок поставки", "Место поставки", "Площадка", "Способ", "Только МСП", "Нацрежим", "ОКПД2", "Группа товара", "Маржа при −10 % (ОСНО)", "Документация (файлы)", "Ссылка"]
ws.append(H)
for r in rows:
    url = f"https://zakupki.gov.ru/epz/order/notice/notice223/common-info.html?regNumber={r['n']}"
    ws.append([r["close"], r["n"], r["cust"], r["title"], r["qty"], r["unit"], r["nmc"], r["per"], r["term"], r["place"], r["etp"], r["method"], r["msp"], r["nat"].strip(), r["okpd"], r["grp"], r["marg"], r["docs"], url])
    X = ws.max_row
    for j in (7, 8): ws.cell(X, j).number_format = "#,##0.00"
    ws.cell(X, 17).number_format = "0%"; ws.cell(X, 19).hyperlink = url
for c in ws[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
for i, w in enumerate([16, 13, 22, 50, 10, 12, 14, 12, 26, 40, 24, 30, 8, 14, 30, 24, 10, 26, 40], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "C2"; ws.auto_filter.ref = f"A1:S{ws.max_row}"
s2 = wb.create_sheet("Спецификация для КП")
s2.append(["Группа товара", "№ п/п", "Наименование (как в закупке)", "Кол-во", "Ед.", "Адрес поставки", "Под закупку №", "Заказчик", "Требуемый срок ответа", "Поставщики для запроса"])
for g, _, sup in GROUPS:
    G = [r for r in rows if r["grp"] == g]
    for i, r in enumerate(G, 1):
        s2.append([g, i, r["name"], r["qty"] if r["qty"] else "уточняется по ТЗ", r["unit"] or "", r["place"], r["n"], r["cust"], "до " + ("09.11" if r["close"] < "2026-11-14" else "12.11"), "; ".join(f"{a} — {b}" for a, b in sup)])
s3 = wb.create_sheet("Поставщики по группам"); s3.append(["Группа", "Поставщик", "Контакт"])
for g, _, sup in GROUPS:
    for a, b in sup: s3.append([g, a, b])
for s in (s2, s3):
    for c in s[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
    for i in range(1, 11): s.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 28
tidy(wb, row_h=30); wb.save("../KONKURSY_KP_0611.xlsx")

# ── Word: запросы КП со спецификацией по группам ──
d = docx.Document(); sec = d.sections[0]; sec.orientation = WD_ORIENT.LANDSCAPE; sec.page_width, sec.page_height = sec.page_height, sec.page_width
for m in ("left_margin", "right_margin", "top_margin", "bottom_margin"): setattr(sec, m, Cm(1.5))
d.styles["Normal"].font.name = "Calibri"; d.styles["Normal"].font.size = Pt(10.5)
d.add_heading("Запросы коммерческих предложений — 06.11.2026", 0)
d.add_paragraph("Шаблон: по одному запросу на группу товара. Отправлять с info@gidropas.ru всем поставщикам группы. Реквизиты в [ ] заполнить.")
for g, _, sup in GROUPS:
    G = [r for r in rows if r["grp"] == g]
    if not G: continue
    d.add_page_break(); d.add_heading(g, 1)
    d.add_paragraph("Кому: " + "; ".join(f"{a} ({b})" for a, b in sup))
    deadline = "09.11.2026" if min(r["close"] for r in G) < "2026-11-14" else "12.11.2026"
    d.add_paragraph(f"Тема: Запрос КП — {g.lower()}, срочно, ответ до {deadline}")
    d.add_paragraph("Добрый день!\n\nООО «ГИДРОПАС» [ИНН ___, г. ___] — поставщик материально-технических ресурсов для предприятий энергетики и ЖКХ. Просим направить коммерческое предложение на поставку продукции по спецификации ниже "
                    f"(для участия в закупках заказчиков). Просим ответить до {deadline}.")
    t = d.add_table(rows=1, cols=6); t.style = "Light Grid Accent 1"
    for i, c in enumerate(["№", "Наименование", "Кол-во", "Ед. изм.", "Адрес поставки", "Требования"]): t.rows[0].cells[i].text = c
    for i, r in enumerate(G, 1):
        c = t.add_row().cells
        c[0].text = str(i); c[1].text = r["name"] or ""; c[2].text = f"{r['qty']:g}" if r["qty"] else "уточняется"; c[3].text = r["unit"] or ""; c[4].text = r["place"] or ""
        c[5].text = "соответствие ГОСТ/ТУ производителя; новая продукция; срок годности ≥ 80 % на дату поставки"
    d.add_paragraph("\nВ КП просим указать:\n"
                    "1) цену за единицу и общую стоимость с НДС 22 % (отдельно — с доставкой до адреса и на условиях самовывоза/FCA);\n"
                    "2) производителя, марку/ТУ/ГОСТ, страну происхождения; номер записи в реестре российской промышленной продукции (ПП 719) или евразийском реестре — при наличии;\n"
                    "3) срок изготовления и поставки (дни с даты заявки) и возможность поставки партиями по графику заказчика в 2027 г.;\n"
                    "4) условия оплаты (возможна ли отсрочка 30–60 дней после поставки);\n"
                    "5) срок действия предложения (не менее 60 дней);\n"
                    "6) комплект документов: паспорт качества, сертификат/декларация, паспорт безопасности, ТУ; при возможности — письмо производителя (авторизация) под закупку;\n"
                    "7) минимальную партию и упаковку.\n\n"
                    "С уважением,\n[ФИО, должность]\nООО «ГИДРОПАС»\ninfo@gidropas.ru, [телефон]")
    d.add_paragraph("Под какие закупки: " + "; ".join(f"№ {r['n']} — {r['cust']}, подача до {r['close'][:10]}" for r in G)).runs[0].italic = True
d.save("../ZAPROSY_KP_0611.docx")
print(len(rows), {g: sum(1 for r in rows if r["grp"] == g) for g, _, _ in GROUPS}, [r["n"] for r in rows if r["grp"] == "Прочее"])
