"""Туркменистан: спрос на фосфонаты/антискаланты/ингибиторы (тендеры 2023–2026, зеркальный импорт Comtrade, сценарии) → ../TURKMENISTAN_HIMIYA_ANALIZ_1111.xlsx + .pdf
Данные: data/tz_lots/turkmenistan_chem/analysis.json (агент 11.11.2026)."""
import json, html, openpyxl
from openpyxl.styles import Font, PatternFill
from pdfutil import html_to_pdf
from tidy import tidy
d = json.load(open("data/tz_lots/turkmenistan_chem/analysis.json"))
T, TS, SC = d["tenders"], d["trade_stats"], d["scenarios"]
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Тендеры"
ws.append(["Заказчик", "№", "Опубликован", "Срок", "Предмет", "Объём", "Допуск", "Взнос", "Релевантность", "Ссылка"])
rank = {"high": 0, "medium": 1, "context": 2, "low": 3}
for t in sorted(T, key=lambda t: rank.get(str(t.get("relevance", "")).split()[0] if t.get("relevance") else "low", 3)):
    ws.append([t.get("customer"), t.get("tender_no"), t.get("published"), t.get("deadline"), t.get("products"), t.get("quantities"), t.get("participation"), t.get("fee"), t.get("relevance"), t.get("url")])
s2 = wb.create_sheet("Импорт (Comtrade)"); s2.append(["HS", "Описание", "Год", "Тонн (без Ирана)", "USD"])
for r in TS["summary_by_hs_year"]: s2.append([r["hs"], r["hs_desc"], r["year"], r.get("mirror_total_excl_Iran_t"), r.get("mirror_total_excl_Iran_usd")])
s3 = wb.create_sheet("Сценарии"); s3.append(["Сценарий", "Рынок, т/год", "ISP 1-й год, т", "ISP 3-й год, т", "Структура", "Допущения"])
for k in ("low", "base", "high"):
    v = SC[k]; s3.append([k, v["market_t"], v["isp_year1_t"], v["isp_year3_t"], v.get("components"), v.get("assumptions")])
s4 = wb.create_sheet("Риски"); s4.append(["Риск"]); [s4.append([r if isinstance(r, str) else json.dumps(r, ensure_ascii=False)]) for r in d["risks"]]
s5 = wb.create_sheet("Конкуренты"); s5.append(["Конкурент"]); [s5.append([c if isinstance(c, str) else json.dumps(c, ensure_ascii=False)]) for c in d["competitors"]]
for s in wb:
    for c in s[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")
for i, w in enumerate([26, 14, 12, 12, 60, 18, 30, 10, 12, 40], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
s2.column_dimensions["B"].width = 60; s3.column_dimensions["E"].width = 40; s3.column_dimensions["F"].width = 70; s4.column_dimensions["A"].width = 140; s5.column_dimensions["A"].width = 140
tidy(wb, row_h=40); wb.save("../TURKMENISTAN_HIMIYA_ANALIZ_1111.xlsx")

e = html.escape
css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:8.4px;margin:0;color:#111}h1{font-size:13px;margin:0 0 4px}h2{font-size:10px;color:#305496;margin:7px 0 3px}
table{border-collapse:collapse;width:100%}th,td{border:1px solid #bbb;padding:2px 4px;vertical-align:top}th{background:#305496;color:#fff;text-align:left}.box{border:1px solid #305496;background:#f3f6fb;padding:4px 7px;margin:3px 0}.bad{background:#f8cbad}.n{text-align:right}"""
yr = {}
for r in TS["summary_by_hs_year"]:
    if r["hs"] in ("293139", "293149"): yr[r["year"]] = r
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body><h1>Туркменистан: рынок фосфонатов, антискалантов и ингибиторов (11.11.2026)</h1>",
 "<div class='box bad'><b>Вывод:</b> 500 т в год — нереалистично. Весь рынок продуктов профиля Kairui оценивается в 260–1 260 т/год (база ≈620 т); прямой импорт фосфонатов 35–195 т/год. "
 "Реалистично для ISP: ≈40 т в 1-й год, ≈120 т на 3-й (база). Кроме того, Туркменгаз (с 09.2025) и Туркменнебит (с 06.2026) допускают к лоту «Химическая продукция» только производителей — заявку должен подавать сам Kairui, ISP — агент. "
 "Условие «минимум 500 т/год» надо перевести в целевой объём (без штрафа; при невыполнении — снятие эксклюзива) или расширить территорию на Центральную Азию.</div>",
 "<h2>1. Импорт фосфонатов (HS 2931.39/2931.49), зеркальные данные UN Comtrade, без Ирана</h2><table><tr><th>Год</th><th>Тонн</th><th>USD</th></tr>"]
for y in sorted(yr): h.append(f"<tr><td>{y}</td><td class=n>{yr[y].get('mirror_total_excl_Iran_t')}</td><td class=n>{(yr[y].get('mirror_total_excl_Iran_usd') or 0):,.0f}</td></tr>".replace(",", " "))
h.append("</table><p>Китай фосфонатов напрямую в 2022–2025 не отгружал; идут смесями (HS 3824.99: ≈7 тыс. т/год всего, Китай 3,7 тыс. т в 2025, но это в основном буровые и прочие добавки). Смолы 3914: ≈109 т (2025). PAC: &lt;20 т/год. Туркменистан сам в Comtrade не отчитывается; Россия — только 2021, ОАЭ/Иран — до 2023 → 2024–2025 занижены.</p>")
h.append("<h2>2. Сценарии (оценка)</h2><table><tr><th>Сценарий</th><th>Рынок, т/год</th><th>ISP 1-й год</th><th>ISP 3-й год</th><th>Допущения</th></tr>")
for k, n in (("low", "Низкий"), ("base", "Базовый"), ("high", "Высокий")):
    v = SC[k]; h.append(f"<tr><td>{n}</td><td class=n>{v['market_t']}</td><td class=n>{v['isp_year1_t']} т</td><td class=n>{v['isp_year3_t']} т</td><td>{e(v.get('assumptions',''))}</td></tr>")
h.append("</table><h2>3. Тендеры: ключевые (из 99 найденных за 2023–2026)</h2><table><tr><th>Заказчик</th><th>№ / дата</th><th>Предмет</th><th>Допуск</th></tr>")
key = [t for t in T if str(t.get("relevance", "")).startswith("high")] + [t for t in T if str(t.get("relevance", "")).startswith("medium")][:7]
for t in key:
    h.append(f"<tr><td>{e(str(t.get('customer')))}</td><td>{e(str(t.get('tender_no') or ''))} {e(str(t.get('published') or ''))}</td><td>{e(str(t.get('products'))[:180])}</td><td>{e(str(t.get('participation') or '')[:90])}</td></tr>")
h.append("</table><p>Лот №5 «Химическая продукция»: Туркменнебит — 38 тендеров (взнос 575 USD), Туркменгаз — 31 (500–1 500 USD). Состав лотов, объёмы и итоги не публикуются. ОЭДФ/НТФ/фосфонаты по названию не встречаются ни разу.</p>"
 "<h2>4. Драйверы и конкуренты</h2><p>Добыча 2025: газ 76,5 млрд м³, нефть 8,38 млн т. Опреснение RO: Кияныли 50 тыс. м³/сут, Аваза 35 тыс. м³/сут, НПЗ Туркменбаши +15 тыс. м³/сут с 2029. Электростанции 6 511 МВт — в основном газотурбинные (мало оборотной воды). "
 "Конкуренты в спецификациях: Hercules, Culligan, Hypersperse, Solisep; импорт — Китай, Турция, Германия, Нидерланды; химия часто идёт пакетом с нефтесервисом. Местный дистрибьютор — ИП Ezber.</p>"
 "<h2>5. Риски</h2><p>Допуск только производителей; привязка спецификаций к западным брендам; взнос 500–1 725 USD за лот и конверт в Ашхабаде; у CNPC окно 7 дней; манат неконвертируем, оплата после поставки → нужен аккредитив; логистика Китай–Хоргос–Балканский велаят 14–15 дней, фосфоновые кислоты — опасный груз класс 8; транзит через Иран исключён.</p>"
 "<h2>6. Что делать</h2><p>1) До подписания дистрибьюторского договора: «500 т» — целевой объём без штрафа, при невыполнении Kairui вправе снять эксклюзив; либо территория «Туркменистан + Казахстан + Узбекистан + Азербайджан». 2) Kairui регистрируется участником тендеров Туркменгаза/Туркменнебита как производитель, ISP — его агент по договору. 3) Прямые поставки: CNPC (ингибиторы коррозии, реагенты для промысловой воды), Eni, опреснители Кияныли/Аваза, Туркменхимия (охлаждение). 4) Местный партнёр для водоканалов.</p>"
 "<p style='font-size:7px;color:#555'>Источники: 145 URL в analysis.json (oilgas.gov.tm, business.com.tm, UN Comtrade public API, сайты операторов). Сценарии — оценки агента с расчётом в файле.</p></body></html>")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/tm.html"
open(hp, "w").write("".join(h)); html_to_pdf(hp, "../TURKMENISTAN_HIMIYA_ANALIZ_1111.pdf"); print("ok")
