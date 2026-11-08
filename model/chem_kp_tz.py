"""Запросы КП заводам под 6 химических лотов по ТЗ (08.11.2026). Спецификации — точно по ТЗ (data/tz_lots/chem_0811/review_*.json),
без заказчиков, номеров закупок и наших цен. Контакты — data/companies/chem_plants_contacts.json.
→ ../ZAPROSY_KP_HIMIYA_TZ/*.docx, ../KONTAKTY_HIMIYA_TZ_0811.xlsx + .pdf, ../PISMA_HIMIYA_TZ_0811.docx"""
import json, os, re, html, docx, openpyxl
from docx.shared import Pt
from openpyxl.styles import Font, PatternFill, Alignment
from pdfutil import html_to_pdf
from tidy import tidy

OUT = "../ZAPROSY_KP_HIMIYA_TZ/"; os.makedirs(OUT, exist_ok=True)
CT = {c["key"]: c for c in json.load(open("data/companies/chem_plants_contacts.json"))}
# позиция: (заголовок, ответ до, строки спецификации)
POS = {
 "natr": ("Натр едкий технический марки РД, высший сорт, ГОСТ Р 55064-2012", "СРОЧНО — до 10.11.2026 (понедельник)", [
   "Массовая доля NaOH — не менее 46 %.",
   "Количество: 75 т в пересчёте на 100 % NaOH (≈163 т раствора 46 %).",
   "Место поставки: г. Салават, Республика Башкортостан.",
   "Срок: отгрузки в декабре 2026 г., каждая партия — не позднее 10 рабочих дней от заявки; вся поставка до 31.12.2026.",
   "Транспорт: автоналив (возможна ж/д — укажите); тара невозвратная или возврат цистерн — укажите условия.",
   "Просим цену за 1 т в пересчёте на 100 % и за 1 т раствора, FCA завод и с доставкой до г. Салавата."]),
 "sa": ("Сульфат алюминия технический очищенный, ВЫСШИЙ сорт, ГОСТ 12966-85", "до 12.11.2026", [
   "Массовая доля Al2O3 — не менее 16 %; Fe2O3 — не более 0,02 %; нерастворимый остаток — не более 0,3 %; As2O3 — не более 0,001 %; размер частиц — не более 20 мм.",
   "Количество: 18 т на 12 месяцев.",
   "Упаковка: мешки по 25 кг на паллетах по 1 т (насыпью нельзя), авто-тент с задней загрузкой.",
   "Место поставки: Тверская обл., Старицкий р-н (очистные сооружения г. Старица).",
   "График: ежемесячно, ≈1,5 т, каждая партия — не позднее 7 календарных дней от заявки.",
   "Обязательно: страна происхождения и номер реестровой записи российской промышленной продукции (ПП 719), паспорт качества на партию."]),
 "soda_vte": ("Сода кальцинированная техническая марка А, 1 сорт, ГОСТ 5100-85", "до 13.11.2026", [
   "Количество: 3 078 т на 2027 г. (поставки равномерно в течение года, ≈250 т/мес.).",
   "Упаковка: мягкие контейнеры (МКР) 800–1 000 кг.",
   "Место поставки: г. Москва, ЮВАО (2-я Вольская ул.), автотранспортом.",
   "Особенность: заявки с коротким сроком (1 рабочий день) — просим условия ответственного хранения/консигнационного склада в Москве или Московской обл.",
   "Просим цену EXW завод (вагон/авто) и с доставкой до Москвы."]),
 "soda_nov": ("Сода кальцинированная техническая марка Б, высший сорт, ГОСТ 5100-85", "до 17.11.2026", [
   "Количество: 843,1 т на 2027 г., ежемесячными партиями (≈70 т/мес.), каждая партия — не позднее 5 календарных дней от заявки.",
   "Упаковка: МКР-1,0 С двухстропные, масса 590–620 кг; доставка автотранспортом с верхней выгрузкой.",
   "Место поставки: г. Великий Новгород.",
   "Обязательно: наименование производителя, страна происхождения, номер реестровой записи российской промышленной продукции."]),
 "soda_ufa": ("Сода кальцинированная техническая марка Б (2 сорт или выше), ГОСТ 5100-85", "до 12.11.2026", [
   "Количество: 50,4 т одной партией в марте 2027 г.; продукция изготовлена не ранее чем за 2 месяца до поставки.",
   "Упаковка: полипропиленовые мешки с плёночным вкладышем / МКР.",
   "Место поставки: г. Уфа, автотранспортом до склада.",
   "Обязательно: номер реестровой записи российской промышленной продукции на дату подачи (до 16.11.2026)."]),
 "gipo": ("Гипохлорит натрия марки А, ГОСТ 11086-76", "до 17.11.2026", [
   "Массовая концентрация активного хлора — не менее 190 г/дм³.",
   "Количество: 607 т на 2027 г. (±25 %).",
   "Тара: канистры 20 л (≈25 кг), ОБОРОТНЫЕ (возврат пустой тары).",
   "Поставка: ≈50 объектов водоснабжения/водоотведения в 15 районах Ленинградской обл., партии от 0,2 т, заявки за 10 дней.",
   "Просим: 1) цену налива EXW завод; 2) цену в канистрах с доставкой по Ленинградской обл., если возможно; 3) есть ли у вас (или партнёров) разлив/склад в СПб/ЛО."]),
}
# адресаты: ключ контакта → позиции (+ копия)
TO = [("1", ["natr", "soda_vte", "soda_nov", "soda_ufa"], "Ilina.TaN@ruschem.ru (отгрузка АО «БСК»)"),
      ("2", ["soda_vte", "soda_nov", "soda_ufa"], "info.bkh@ruschem.ru (ТД «Башхим» — по данным БСК соду БСЗ продаёт ТД)"),
      ("3", ["natr", "soda_vte", "soda_nov", "soda_ufa"], None),
      ("4", ["natr", "gipo"], "salova@kaustik.ru (приёмная коммерческого директора)"),
      ("5", ["natr", "gipo"], "himprom@himprom.com"),
      ("7", ["natr"], None),
      ("12c", ["sa"], None), ("12a", ["sa"], "tehotdel@sintez-k.ru (технический отдел)"), ("9", ["sa", "gipo"], None),
      ("8c", ["gipo"], None), ("8a", ["gipo"], "info@halopolymer-td.com (ТД «ГалоПолимер»)")]

def text(ps):
    t = ["Добрый день!", "", "ООО «ГИДРОПАС» (Москва) — поставщик химических реагентов для предприятий энергетики и водоканалов — просит направить коммерческое предложение на поставку:", ""]
    for p in ps:
        title, due, rows = POS[p]
        t.append(f"■ {title} — ответ {due}")
        t += [f"   • {r}" for r in rows]; t.append("")
    t += ["По каждой позиции просим указать:",
          "1) цену за единицу с НДС (EXW/FCA завод и с доставкой до места поставки);",
          "2) минимальную партию, тару, срок отгрузки от заявки;",
          "3) условия оплаты (возможна ли отсрочка 15–30 дней);",
          "4) срок действия цены и возможность фиксации цены на 2027 г.;",
          "5) паспорт качества / сертификат, страну происхождения и номер записи в реестре российской промышленной продукции (ПП 719);",
          "6) возможность заключить прямой договор / выдать письмо поддержки на ООО «ГИДРОПАС».", "",
          "С уважением,", "[ФИО, должность]", "ООО «ГИДРОПАС», ИНН [___]", "info@gidropas.ru, [телефон]"]
    return "\n".join(t)

rows = []
for k, ps, cc in TO:
    c = CT[k]
    SH = {"natr": "натр едкий РД", "sa": "сульфат алюминия высший сорт", "soda_vte": "сода кальцинированная", "soda_nov": "сода кальцинированная", "soda_ufa": "сода кальцинированная", "gipo": "гипохлорит натрия марки А"}
    subj = "Запрос КП — " + ", ".join(sorted({SH[p] for p in ps})) + " — ООО «ГИДРОПАС»"
    if "natr" in ps: subj = "СРОЧНО: " + subj
    body = text(ps)
    d = docx.Document(); d.styles["Normal"].font.name = "Calibri"; d.styles["Normal"].font.size = Pt(11)
    d.add_paragraph(f"Кому: {c['name']}\nE-mail: {c.get('email') or 'не найден — звонить'}" + (f"\nКопия: {cc}" if cc else "") + f"\nТел.: {c.get('phone') or '—'}")
    d.add_paragraph("Тема: " + subj).runs[0].bold = True
    for para in body.split("\n"): d.add_paragraph(para)
    fn = re.sub(r"[^\w\-]+", "_", f"{k}_{c['name']}")[:70] + ".docx"; d.save(OUT + fn)
    due = min((POS[p][1] for p in ps), key=lambda s: ("СРОЧНО" not in s, s[-10:]))
    rows.append(dict(name=c["name"], inn=c.get("inn"), email=c.get("email"), cc=cc, phone=c.get("phone"), person=c.get("person"), st="да" if c.get("verified") else "только сайт",
                     lots=", ".join(POS[p][0].split(",")[0] for p in ps), due=due, subj=subj, body=body, fn=fn, note=c.get("note")))
rows.sort(key=lambda r: ("СРОЧНО" not in r["due"], r["due"][-10:]))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Кому отправить"
ws.append(["№", "Компания", "E-mail (Кому)", "Копия", "Телефон", "Контактное лицо", "Что запрашиваем", "Ответ до", "Контакт подтверждён", "Тема письма", "Текст письма (копировать)", "Файл (ZAPROSY_KP_HIMIYA_TZ/)", "Примечание"])
for i, r in enumerate(rows, 1):
    ws.append([i, r["name"], r["email"] or "НЕТ — звонить", r["cc"], r["phone"], r["person"], r["lots"], r["due"], r["st"], r["subj"], r["body"], r["fn"], r["note"]])
    ws.cell(ws.max_row, 9).fill = PatternFill("solid", fgColor="c6efce" if r["st"] == "да" else "fff2cc")
for c in ws[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
for i, w in enumerate([4, 36, 30, 34, 26, 26, 40, 18, 11, 50, 100, 40, 50], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
tidy(wb, row_h=60)
for row in ws.iter_rows(min_row=2): row[10].alignment = Alignment(wrap_text=True, vertical="top"); ws.row_dimensions[row[0].row].height = 380
wb.save("../KONTAKTY_HIMIYA_TZ_0811.xlsx")

P = docx.Document(); P.styles["Normal"].font.name = "Calibri"; P.styles["Normal"].font.size = Pt(10.5)
P.add_heading("Запросы КП заводам по ТЗ химических лотов — 08.11.2026", 0)
P.add_paragraph("Отправлять с info@gidropas.ru. Сначала — «СРОЧНО» (натр, ответ до 10.11). Реквизиты в [ ] заполнить. В письмах нет заказчиков, номеров закупок и наших цен.")
for r in rows:
    P.add_page_break(); P.add_heading(r["name"], 1)
    P.add_paragraph(f"Кому: {r['email'] or 'e-mail не найден — звонить'}" + (f"\nКопия: {r['cc']}" if r["cc"] else "") + f"\nТелефон: {r['phone'] or '—'}" + (f"\nКонтакт: {r['person']}" if r["person"] else ""))
    P.add_paragraph("Тема: " + r["subj"]).runs[0].bold = True
    for para in r["body"].split("\n"): P.add_paragraph(para)
P.save("../PISMA_HIMIYA_TZ_0811.docx")

e = lambda s: html.escape(str(s if s not in (None, "") else "—"))
css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:7.6px;color:#111;margin:0}h1{font-size:12px;margin:0 0 3px}h2{font-size:10px;margin:5px 0 2px;border-bottom:2px solid #305496}table{border-collapse:collapse;width:100%}
th,td{border:1px solid #bbb;padding:2px 3px;vertical-align:top}th{background:#305496;color:#fff;text-align:left}.ok{background:#c6efce}.warn{background:#fff2cc}.box{border:1px solid #305496;padding:4px 6px;background:#f3f6fb;margin:3px 0}.s{font-size:6.8px;color:#444}"""
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body><h1>Химия по ТЗ: кому отправить запросы КП ({len(rows)} писем), 08.11.2026</h1>",
     "<div class='box'>Отправлять с info@gidropas.ru. Тексты — в Excel (колонка K) и PISMA_HIMIYA_TZ_0811.docx. Спецификации — точно по ТЗ, без заказчиков, номеров закупок и наших цен. "
     "Башкирская содовая компания: настоящие адреса только @ruschem.ru (sodastr@soda.ru — канцелярия). ЕТК — e-mail не найден, звонить.</div>",
     "<table><tr><th>№</th><th>Компания</th><th>Кому</th><th>Копия</th><th>Телефон</th><th>Что запрашиваем</th><th>Ответ до</th><th>Подтв.</th></tr>"]
for i, r in enumerate(rows, 1):
    h.append(f"<tr><td>{i}</td><td>{e(r['name'])}</td><td><b>{e(r['email'] or 'НЕТ — звонить')}</b></td><td class=s>{e(r['cc'])}</td><td class=s>{e(r['phone'])}</td><td class=s>{e(r['lots'])}</td><td>{e(r['due'])}</td><td class={'ok' if r['st']=='да' else 'warn'}>{e(r['st'])}</td></tr>")
h.append("</table><h2>Что указано в письмах (по ТЗ)</h2><table><tr><th>Товар</th><th>Требования</th><th>Ответ до</th></tr>")
for p, (t, due, rr) in POS.items(): h.append(f"<tr><td><b>{e(t)}</b></td><td class=s>{'<br>'.join(e(x) for x in rr)}</td><td>{e(due)}</td></tr>")
h.append("</table></body></html>")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/chem_kp_tz.html"
open(hp, "w").write("".join(h)); html_to_pdf(hp, "../KONTAKTY_HIMIYA_TZ_0811.pdf")
for r in rows: print(r["due"][:22], "|", r["name"][:40], "|", r["email"], "|", r["subj"][:90])
