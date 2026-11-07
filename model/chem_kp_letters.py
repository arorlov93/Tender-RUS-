"""Запросы КП заводам/трейдерам химии по маржинальным лотам (07.11.2026).
data/companies/chem_plants_contacts.json (контакты, проверка по 2 источникам) → ../ZAPROSY_KP_HIMIYA/*.docx, ../KONTAKTY_HIMIYA_ZAVODY_0711.xlsx + .pdf, ../PISMA_HIMIYA_ZAVODY_0711.docx
В письмах НЕ указываем заказчиков, номера закупок и наши цены — только товар, объём, город, период (чтобы завод не зашёл в торги напрямую)."""
import json, os, re, html, docx, openpyxl
from docx.shared import Pt
from openpyxl.styles import Font, PatternFill, Alignment
from pdfutil import html_to_pdf
from tidy import tidy

OUT = "../ZAPROSY_KP_HIMIYA/"; os.makedirs(OUT, exist_ok=True)
C = json.load(open("data/companies/chem_plants_contacts.json"))
# товар → позиции (город, объём, период/условия) — из открытых лотов, без заказчиков и цен
G = {
 "soda": ("Сода кальцинированная техническая марка Б, высший сорт, ГОСТ 5100-85", "до 11.11.2026", [
     ("Москва (юго-восток)", "3 078 т", "2027 г., партиями по графику; тара — навалом/МКР (просим варианты)"),
     ("Великий Новгород", "834 т", "2027 г., партиями"),
     ("Уфа", "50,4 т", "2027 г.")]),
 "natr": ("Натр едкий технический (каустическая сода) марки РД, высший сорт, ГОСТ Р 55064-2012", "до 11.11.2026", [
     ("Республика Татарстан (Казань)", "75 т", "2027 г.; цену просим за т раствора и за т в пересчёте на 100 % NaOH"),
     ("ДНР (г. Зугрэс)", "2 511 т", "2027 г., ж/д цистернами; просим цену FCA станция завода и с доставкой")]),
 "gipo": ("Гипохлорит натрия марки А, ГОСТ 11086-76", "до 12.11.2026", [
     ("Ленинградская область (несколько адресов)", "607 т", "2027 г., в канистрах"),
     ("Московская область", "300 т", "2027 г., налив/кубы — просим оба варианта"),
     ("Ленинградская область", "70 т", "2027 г."),
     ("Краснодарский край (Геленджик)", "≈56 т", "2027 г.")]),
 "tio": ("Натрия тиосульфат кристаллический (5-водный), ГОСТ 244-76", "СРОЧНО — до 10.11.2026", [("Саратов", "10 т", "2027 г., мешки")]),
 "ntf": ("Кислота нитрилотриметилфосфоновая (НТФ, ATMP) 95 %, порошок", "СРОЧНО — до 11.11.2026", [("Владимирская область", "14 т", "2027 г., мешки 25 кг")]),
 "sa": ("Сульфат алюминия технический очищенный, ГОСТ 12966-85", "до 12.11.2026", [
     ("Тверь", "18 т (твёрдый)", "2027 г.; марка/сорт уточняются — просим цены высшего и 1 сорта"),
     ("Ярославль", "6 000 т (водный раствор)", "2027 г., автоцистернами; просим цену за т раствора и концентрацию Al2O3")]),
 "tnf": ("Тринатрийфосфат технический, ГОСТ 201-76", "до 11.11.2026", [("Набережные Челны", "≈14 т", "2027 г., мешки")]),
}
EN = {"ntf": ("ATMP (amino trimethylene phosphonic acid) 95 % powder", [("Vladimir region, Russia", "14 t", "2027, 25 kg bags")])}
# компания (номер в списке агента) → товары
MAP = {"1": ["soda", "natr"], "2": ["soda"], "3": ["soda", "natr", "tio"], "4": ["natr", "gipo"], "5": ["natr", "gipo"], "6": ["natr"], "7": ["natr"], "8": ["gipo"], "9": ["gipo", "ntf", "sa", "tio"],
       "10": ["tio"], "11": ["ntf"], "12": ["sa"], "13": ["tnf"]}
SKIP = {"10a", "12b", "13a"}  # 10a — дубль Аквахим (key 9); БИАКР и Ветлужский ХЗ — контакты не найдены (ТНФ — через ТД «Химспециализация»)
def prods(c):
    if str(c.get("key")) in SKIP: return []
    k = re.match(r"\d+", str(c.get("key")) or "")
    return MAP.get(k.group(0) if k else "", [])

def body_ru(name, ps):
    parts = ["Добрый день!", "",
             f"ООО «ГИДРОПАС» (Москва) — поставщик химических реагентов для предприятий энергетики и водоканалов — формирует закупки на 2027 год и просит направить коммерческое предложение на следующую продукцию:", ""]
    for p in ps:
        title, due, rows = G[p]
        parts.append(f"■ {title} (ответ {due}):")
        for city, q, cond in rows: parts.append(f"   • {city} — {q}; {cond}")
        parts.append("")
    parts += ["Просим указать:",
              "1) цену за тонну с НДС — EXW/FCA завод (станция отгрузки) и, если возможно, с доставкой до указанных городов;",
              "2) минимальную партию (вагон/цистерна/автомашина), тару, срок отгрузки после заявки;",
              "3) условия оплаты (возможна ли отсрочка 15–30 дней или частичная предоплата);",
              "4) срок действия цены и порядок её изменения в течение 2027 г. (договор на год с графиком поставок);",
              "5) паспорт качества / сертификат соответствия, наличие записи в реестре российской промышленной продукции (ПП 719), страну происхождения;",
              "6) возможность заключить прямой договор поставки или дилерское соглашение с ООО «ГИДРОПАС».", "",
              "Объёмы ориентировочные, итоговые количества и график сообщим после согласования. Будем благодарны за оперативный ответ.", "",
              "С уважением,", "[ФИО, должность]", "ООО «ГИДРОПАС», ИНН [___]", "info@gidropas.ru, [телефон]"]
    return "\n".join(parts)
def body_en(ps):
    parts = ["Dear Sirs,", "", "GIDROPAS LLC (Moscow, Russia), a supplier of water-treatment chemicals to power plants and water utilities, requests your quotation for:", ""]
    for p in ps:
        t, rows = EN[p]
        parts.append(f"■ {t}:")
        for city, q, cond in rows: parts.append(f"   • {city} — {q}; {cond}")
    parts += ["", "Please advise: price per tonne FOB China port and, if possible, CIF St. Petersburg / DAP Moscow; minimum order; packaging; lead time; payment terms (USD/CNY); COA of a recent batch (active content, chloride, phosphorous acid); validity of the offer.",
              "Reply by 11.11.2026 would be highly appreciated.", "", "Best regards,", "GIDROPAS LLC", "info@gidropas.ru, [phone]"]
    return "\n".join(parts)

rows = []
for c in C:
    ps = prods(c)
    if not ps: continue
    en = c.get("lang") == "en"
    ps_en = [p for p in ps if p in EN] if en else []
    if en and not ps_en: continue
    subj = ("RFQ — " + ", ".join(EN[p][0].split(" (")[0] for p in ps_en) + " — reply by 11.11.2026") if en else \
           ("Запрос КП на 2027 г. — " + ", ".join(G[p][0].split(",")[0].split(" (")[0] for p in ps) + " — ООО «ГИДРОПАС»")
    text = body_en(ps_en) if en else body_ru(c.get("name"), ps)
    d = docx.Document(); d.styles["Normal"].font.name = "Calibri"; d.styles["Normal"].font.size = Pt(11)
    d.add_paragraph(f"Кому: {c.get('name')}\nE-mail: {c.get('email') or '[не найден — звонить]'}   Тел.: {c.get('phone') or '—'}")
    d.add_paragraph("Тема: " + subj).runs[0].bold = True
    for para in text.split("\n"): d.add_paragraph(para)
    fn = re.sub(r"[^\w\-]+", "_", f"{c.get('key')}_{c.get('name')}")[:70] + ".docx"
    d.save(OUT + fn)
    st = "да" if c.get("verified") else ("частично" if re.search(r"2 источник|оба|два источник|\+", c.get("sources") or "", re.I) else "нет")
    first = min((G[p][1] for p in ps), key=lambda s: ("СРОЧНО" not in s, s))
    rows.append(dict(key=c.get("key"), name=c.get("name"), inn=c.get("inn"), prod=", ".join(G[p][0].split(",")[0] for p in ps), email=c.get("email"), phone=c.get("phone"), person=c.get("person"),
                     site=c.get("site"), st=st, src=(c.get("sources") or "") + ("; " + c["note"] if c.get("note") else ""), subj=subj, text=text, fn=fn, due=first, lang="EN" if en else "RU"))
rows.sort(key=lambda r: ("СРОЧНО" not in r["due"], r["prod"]))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Кому писать"
ws.append(["№", "Компания", "ИНН", "Что запрашиваем", "Ответ до", "E-mail", "Телефон", "Контактное лицо", "Сайт", "Контакт подтверждён (2 источника)", "Источники / примечание", "Язык", "Тема письма", "Текст письма (копировать в e-mail)", "Файл письма (ZAPROSY_KP_HIMIYA/)"])
for i, r in enumerate(rows, 1):
    ws.append([i, r["name"], r["inn"], r["prod"], r["due"], r["email"], r["phone"], r["person"], r["site"], r["st"], r["src"], r["lang"], r["subj"], r["text"], r["fn"]])
    ws.cell(ws.max_row, 10).fill = PatternFill("solid", fgColor={"да": "c6efce", "частично": "fff2cc"}.get(r["st"], "f8cbad"))
s2 = wb.create_sheet("Объёмы в письмах"); s2.append(["Товар", "Город", "Объём", "Условия", "Ответ до"])
for p, (t, due, rr) in G.items():
    for city, q, cond in rr: s2.append([t, city, q, cond, due])
for s in wb:
    for c in s[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
for i, w in enumerate([4, 36, 13, 30, 16, 34, 28, 24, 22, 12, 50, 6, 50, 100, 40], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
for i, w in enumerate([50, 34, 16, 60, 18], 1): s2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
tidy(wb, row_h=60)
for row in ws.iter_rows(min_row=2):
    row[13].alignment = Alignment(wrap_text=True, vertical="top"); ws.row_dimensions[row[0].row].height = 300
wb.save("../KONTAKTY_HIMIYA_ZAVODY_0711.xlsx")

P = docx.Document(); P.styles["Normal"].font.name = "Calibri"; P.styles["Normal"].font.size = Pt(10.5)
P.add_heading("Запросы КП заводам и трейдерам химии — 07.11.2026", 0)
P.add_paragraph("Отправлять с info@gidropas.ru. В письмах нет заказчиков, номеров закупок и наших цен. Реквизиты в [ ] заполнить. Сначала — помеченные «СРОЧНО».")
for r in rows:
    P.add_page_break(); P.add_heading(r["name"] or "—", 1)
    P.add_paragraph(f"Что запрашиваем: {r['prod']} · ответ {r['due']} · контакт подтверждён: {r['st']}")
    P.add_paragraph(f"Кому (e-mail): {r['email'] or '— не найден, звонить'}\nТелефон: {r['phone'] or '—'}" + (f"\nКонтактное лицо: {r['person']}" if r["person"] else ""))
    P.add_paragraph("Тема: " + r["subj"]).runs[0].bold = True
    for para in r["text"].split("\n"): P.add_paragraph(para)
P.save("../PISMA_HIMIYA_ZAVODY_0711.docx")

e = lambda s: html.escape(str(s if s not in (None, "") else "—"))
css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:7.4px;color:#111;margin:0}h1{font-size:12px;margin:0 0 3px}h2{font-size:10px;margin:5px 0 2px;border-bottom:2px solid #305496}table{border-collapse:collapse;width:100%}
th,td{border:1px solid #bbb;padding:1px 3px;vertical-align:top}th{background:#305496;color:#fff;text-align:left}.ok{background:#c6efce}.warn{background:#fff2cc}.bad{background:#f8cbad}.box{border:1px solid #305496;padding:4px 6px;background:#f3f6fb;margin:3px 0}.s{font-size:6.6px;color:#444}"""
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body><h1>Химия: кому писать — заводы и трейдеры ({len(rows)} адресатов), 07.11.2026</h1>",
     "<div class='box'>Цель — цены вагонных/крупных партий (уровень завода) по 7 товарам из маржинальных лотов. В письмах нет заказчиков, номеров закупок и наших цен — только товар, объём, город, 2027 г. "
     "Зелёный — e-mail и телефон подтверждены 2 источниками; жёлтый — частично; красный — 1 источник (перед отправкой позвонить). Полные тексты — в Excel (колонка N) и PISMA_HIMIYA_ZAVODY_0711.docx.</div>",
     "<table><tr><th>№</th><th>Компания</th><th>Что запрашиваем</th><th>Ответ до</th><th>E-mail</th><th>Телефон</th><th>2 ист.</th><th>Тема</th></tr>"]
for i, r in enumerate(rows, 1):
    c = {"да": "ok", "частично": "warn"}.get(r["st"], "bad")
    h.append(f"<tr><td>{i}</td><td>{e(r['name'])}<br><span class=s>ИНН {e(r['inn'])}</span></td><td>{e(r['prod'])}</td><td>{e(r['due'])}</td><td>{e(str(r['email'] or '')[:70])}</td><td>{e(str(r['phone'] or '')[:45])}</td><td class={c}>{e(r['st'])}</td><td class=s>{e(r['subj'])}</td></tr>")
h.append("</table><h2>Объёмы, которые указаны в письмах</h2><table><tr><th>Товар</th><th>Город</th><th>Объём</th><th>Условия</th></tr>")
for p, (t, due, rr) in G.items():
    for city, q, cond in rr: h.append(f"<tr><td>{e(t)}</td><td>{e(city)}</td><td>{e(q)}</td><td class=s>{e(cond)}</td></tr>")
h.append("</table></body></html>")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/chem_kp.html"
open(hp, "w").write("".join(h)); html_to_pdf(hp, "../KONTAKTY_HIMIYA_ZAVODY_0711.pdf")
print(len(rows), sum(r["st"] == "да" for r in rows), sum(r["st"] == "частично" for r in rows))
