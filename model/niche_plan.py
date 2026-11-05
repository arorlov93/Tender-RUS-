"""Частота конкурсов ниши (data/companies/niche_frequency.json) → ../NISHA_CHASTOTA.pdf/.xlsx; план развития компании-оптовика → ../PLAN_RAZVITIYA_KOMPANII.pdf"""
import json, html, ast, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from pdfutil import html_to_pdf
from tidy import tidy

d = json.load(open("data/companies/niche_frequency.json"))
def P(v):
    if isinstance(v, str):
        try: return ast.literal_eval(v)
        except Exception:
            try: return float(v)
            except Exception: return v
    return v
SG = [{k: P(v) for k, v in s.items()} for s in d["segments"]]
TU = d["totals_unique"]
e = lambda s: html.escape(str(s if s is not None else "—"))
fm = lambda v: f"{v:,.0f}".replace(",", " ") if isinstance(v, (int, float)) else "—"
mln = lambda v: f"{v/1e6:,.1f}".replace(",", " ") if isinstance(v, (int, float)) else "—"
pc = lambda v: f"{v*100:.0f} %" if isinstance(v, (int, float)) else "—"
css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:8.4px;color:#111;margin:0;line-height:1.28}h1{font-size:14px;margin:0 0 3px}h2{font-size:10.8px;margin:7px 0 3px;border-bottom:2px solid #305496;color:#1f3864}
table{border-collapse:collapse;width:100%;margin:2px 0 4px}th,td{border:1px solid #bbb;padding:1px 3px;vertical-align:top}th{background:#305496;color:#fff;text-align:left}td.n{text-align:right;white-space:nowrap}
.small{font-size:7.2px;color:#444}.box{border:1px solid #305496;padding:4px 6px;background:#f3f6fb;margin:4px 0}ul{margin:2px 0 3px 15px;padding:0}.ok{background:#c6efce}.warn{background:#fff2cc}"""
MON = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]

# ─── частота ───
f223 = [s for s in SG if s["law"] == "fz223"]; f44 = [s for s in SG if s["law"] == "fz44"]
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body><h1>Частота и объём конкурсов ниши (реагенты, иониты, угли, адсорбенты, катализаторы, гликоли, запчасти компрессоров)</h1>",
     "<div class='small'>ЕИС через gosplan: 223-ФЗ с 10.2024, 44-ФЗ с 01.2024; последние 12 мес. = 11.2025–10.2026; без закупок у единственного поставщика и дублей. Учтены лоты, где предмет назван в заголовке — «МТР»/«химпродукция» и закрытые площадки нефтяников не попали, поэтому это нижняя оценка (катализаторы и нефтепромысловая химия занижены в разы).</div>",
     f"<div class='box'><b>Рынок за 12 мес.:</b> 223-ФЗ — {TU['fz223']['last12m_count']} закупок на {mln(TU['fz223']['last12m_nmc_sum'])} млн; в целевом чеке 0,5–30 млн — <b>{TU['fz223']['last12m_count_0_5_30m']} закупок на {mln(TU['fz223']['last12m_nmc_sum_0_5_30m'])} млн</b>, из них {TU['fz223']['last12m_msp_only_0_5_30m']} только для МСП. "
     f"44-ФЗ — {TU['fz44']['last12m_count']} закупок на {mln(TU['fz44']['last12m_nmc_sum'])} млн (мелкие бюджетные лоты). <b>Итого ≈{TU['fz223']['last12m_count']+TU['fz44']['last12m_count']} конкурсов в год, ≈{TU['fz223']['last12m_count_0_5_30m']+TU['fz44']['last12m_count_0_5_30m']} в целевом чеке, доступно ≈{(TU['fz223']['last12m_nmc_sum_0_5_30m']+TU['fz44']['last12m_nmc_sum_0_5_30m'])/1e9:.1f} млрд ₽/год НМЦ.</b> "
     "Это ≈65 подходящих конкурсов в месяц — поток достаточный для компании из 3–5 человек.</div>",
     "<h2>По сегментам, 223-ФЗ (последние 12 мес.)</h2><table><tr><th>Сегмент</th><th>Закупок</th><th>Σ НМЦ, млн</th><th>Медиана НМЦ, млн</th><th>В чеке 0,5–30 млн</th><th>Σ в чеке, млн</th><th>Только МСП</th><th>Сезонность</th></tr>"]
for s in sorted(f223, key=lambda s: -(s.get("nmc_sum_0_5_30m_last12m") or 0)):
    h.append(f"<tr><td>{e(s['segment'])}</td><td class=n>{fm(s.get('count_last12m'))}</td><td class=n>{mln(s.get('nmc_sum_last12m'))}</td><td class=n>{mln(s.get('nmc_median'))}</td><td class=n>{fm(s.get('count_0_5_30m_last12m'))}</td><td class=n>{mln(s.get('nmc_sum_0_5_30m_last12m'))}</td><td class=n>{pc(s.get('share_msp'))}</td><td class=small>{e(s.get('seasonality'))[:150]}</td></tr>")
h.append("</table><h2>По месяцам (223-ФЗ, число закупок, все сегменты)</h2>")
months = sorted({m for s in f223 for m, v in (s.get("by_month") or {}).items() if v is not None})
tot = {m: sum(((s.get("by_month") or {}).get(m) or 0) for s in f223) for m in months}
mx = max(tot.values()) or 1
h.append("<table><tr>" + "".join(f"<th style='font-size:6.5px'>{m[2:]}</th>" for m in months) + "</tr><tr>" +
         "".join(f"<td style='vertical-align:bottom;height:60px;text-align:center'><div style='background:#305496;height:{int(55*tot[m]/mx)}px'></div><span style='font-size:6.5px'>{tot[m]}</span></td>" for m in months) + "</tr></table>")
h.append("<h2>Топ-заказчики (223-ФЗ, все сегменты)</h2><table><tr><th>Сегмент</th><th>Заказчики (число закупок / сумма, млн)</th></tr>")
for s in f223:
    tc = s.get("top_customers") or []
    h.append(f"<tr><td>{e(s['segment'])}</td><td class=small>{'; '.join(e(c.get('name'))[:45] + f' ({c.get(chr(99)+chr(111)+chr(117)+chr(110)+chr(116))} / {mln(c.get(chr(115)+chr(117)+chr(109)))})' for c in tc[:7])}</td></tr>")
h.append("</table><h2>44-ФЗ (последние 12 мес.)</h2><table><tr><th>Сегмент</th><th>Закупок</th><th>Σ НМЦ, млн</th><th>Медиана, млн</th></tr>")
for s in f44: h.append(f"<tr><td>{e(s['segment'])}</td><td class=n>{fm(s.get('count_last12m'))}</td><td class=n>{mln(s.get('nmc_sum_last12m'))}</td><td class=n>{mln(s.get('nmc_median'))}</td></tr>")
h.append("</table></body></html>")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/chast.html"
open(hp, "w").write("".join(h)); html_to_pdf(hp, "../NISHA_CHASTOTA.pdf")

wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Сегменты"
cols = ["segment", "law", "count_last12m", "nmc_sum_last12m", "nmc_median", "nmc_p25", "nmc_p75", "count_0_5_30m_last12m", "nmc_sum_0_5_30m_last12m", "share_msp", "seasonality"]
ws.append(["Сегмент", "Закон", "Закупок 12 мес.", "Σ НМЦ 12 мес.", "Медиана НМЦ", "P25", "P75", "В чеке 0,5–30 млн (12 мес.)", "Σ в чеке", "Доля «только МСП»", "Сезонность"])
for s in SG: ws.append([s.get(c) if not isinstance(s.get(c), (dict, list)) else str(s.get(c)) for c in cols])
for row in ws.iter_rows(min_row=2):
    for j in (3, 4, 5, 6, 7, 8): row[j].number_format = "#,##0"
    row[9].number_format = "0%"
s2 = wb.create_sheet("По месяцам"); s2.append(["Сегмент", "Закон"] + months)
for s in SG: s2.append([s["segment"], s["law"]] + [(s.get("by_month") or {}).get(m) for m in months])
s3 = wb.create_sheet("Топ-заказчики"); s3.append(["Сегмент", "Закон", "Заказчик", "ИНН", "Закупок", "Сумма"])
for s in SG:
    for c in s.get("top_customers") or []: s3.append([s["segment"], s["law"], c.get("name"), c.get("inn"), c.get("count"), c.get("sum")])
s4 = wb.create_sheet("Примеры"); s4.append(["Сегмент", "Закон", "Пример"])
for s in SG:
    for x in s.get("examples") or []: s4.append([s["segment"], s["law"], str(x)])
s5 = wb.create_sheet("Методика"); s5.append(["Методика", str(d["method"])[:30000]]); s5.append(["Итог", str(d["summary"])[:30000]])
for s in wb:
    for c in s[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
    for i in range(1, 12): s.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 22
tidy(wb, row_h=20); wb.save("../NISHA_CHASTOTA.xlsx")

# ─── план развития ───
avg = TU["fz223"]["last12m_nmc_sum_0_5_30m"] / TU["fz223"]["last12m_count_0_5_30m"]   # средний лот в чеке
# маржа — из 8 посчитанных лотов (niche_margin.json), без 2 лотов низкой уверенности; зависит от снижения цены на торгах
MT = [t for t in json.load(open("data/companies/niche_margin.json"))["tenders"] if t.get("margin_m10") is not None and t.get("confidence") != "низкая"]
mg = lambda red: 1 - sum(t["cost"]["total_no_vat"] for t in MT) / sum(t["nmc_no_vat"] * (1 - red) for t in MT)
SCEN = [("Осторожный: снижение −30 %", .30), ("Базовый: снижение −20 %", .20), ("Оптимистичный: снижение −10 %", .10)]
YEARS = [("Год 1 (старт, ЕИС-энергетика)", 150, .12), ("Год 2 (+ дилерства, вся водоподготовка)", 300, .15), ("Год 3 (+ импорт из Китая, начало нефтегаза)", 450, .17)]
SC = []
for yn, n, w in YEARS:
    for sn, red in SCEN:
        rev = n * w * avg * (1 - red) / 1.22
        SC.append((yn, sn, n, w, red, mg(red), rev, rev * mg(red)))
S = ["<h1>План развития компании-оптовика (реагенты, иониты, угли, адсорбенты → катализаторы)</h1>",
     f"<div class='small'>Опирается на: частоту конкурсов ЕИС за 12 мес. (≈780 лотов/год в чеке 0,5–30 млн, ≈3,3 млрд НМЦ), маржу по 51 текущему лоту (по 8 посчитанным — 26–74 % при −10 %), модели СПб НХК (1,4–2,0 млрд выручки) и БПО Евразия-Снаб (0,7 млрд). Сценарий — оценочный, допущения указаны.</div>",
     "<div class='box'><b>Стратегия:</b> зайти через простой и частый товар энергетики (иониты, реагенты водоподготовки ТЭЦ, угли, силикагель/цеолит для сетей, гликоль) на открытых ЕИС-лотах, набрать референсы и дилерства, затем выйти на закрытые площадки нефтегаза и в катализаторы — там чеки больше, а конкурентов меньше.</div>",
     "<h2>1. Этапы</h2><table><tr><th>Этап</th><th>Что делаем</th><th>Результат этапа</th></tr>"]
for a, b, c in [
    ("0–3 мес.: запуск", "ООО на ОСНО, МСП; аккредитация на 6–8 площадках; 5–8 дилерских договоров (СКАТЗ, KNT, «Сорбент», производители ионитов и реагентов ТЭЦ, 2–3 китайских); сайт; тендерный специалист + технолог по договору. Подача 40–60 заявок в лоты иониты/угли/реагенты ТЭЦ/силикагель (Т Плюс, ТГК-1/2, Форвард, Мосэнерго, Интер РАО, Россети).", "5–8 договоров, первые референсы, отлаженная логистика"),
    ("3–12 мес.: рост в энергетике", "Расширение до всех сегментов водоподготовки и ионитов по РФ; 150+ заявок/год; склад ходовых позиций в Московском регионе; прямой импорт ионитов/углей/гидразина из Китая (маржа выше, чем у российских дилеров); предквалификация в Газпром нефти, ЛУКОЙЛе, СИБУРе, НОВАТЭКе.", "15–25 договоров, допуски к закрытым площадкам"),
    ("12–24 мес.: нефтегаз", "Реагенты нефтедобычи (деэмульгаторы, ингибиторы, поглотители H₂S) через опытно-промысловые испытания у Русвьетпетро, Оренбургнефтеотдачи, Ульяновскнефтегаза; адсорбенты для газопереработки; штатный технолог; аттестация продукции в Россетях вместе с производителями.", "Выход в чеки 30–150 млн, первые контракты на закрытых площадках"),
    ("24–36 мес.: катализаторы и сервис", "Катализаторы НПЗ и нефтехимии (российские заводы KNT/СКАТЗ/НИАП + китайские), шеф-монтаж загрузки через партнёров, регистрация в ГИИС ДМДК для драгметаллсодержащих; рамочные договоры на 2–3 года.", "Модель СПб НХК: 0,5–1,5 млрд выручки")]:
    S.append(f"<tr><td><b>{a}</b></td><td>{b}</td><td>{c}</td></tr>")
S.append("</table><h2>2. Финансовый сценарий — маржа из реального расчёта лотов</h2>")
S.append(f"<div class='box'><b>Откуда маржа.</b> По 6 лотам средней уверенности (угли Т Плюс Владимир и Саратов, гидразин Т Плюс, Silcarbon Форвард, иониты ШТЭС и ГК СЭТ) себестоимость = {sum(t['cost']['total_no_vat'] for t in MT)/sum(t['nmc_no_vat'] for t in MT)*100:.0f} % НМЦ без НДС. "
         f"Отсюда маржа: при победе с −10 % — {mg(.1)*100:.0f} %, с −20 % — {mg(.2)*100:.0f} %, с −30 % — {mg(.3)*100:.0f} %. Гликоль ТГК-1 и ингибитор Форварда (74 % и 58 %) исключены — низкая уверенность. "
         "Насколько придётся снижать, решает конкуренция: в товарных лотах (угли, иониты, гидразин) дилеры обычно дают −15…−30 %, в брендовых реагентах — почти без снижения. Поэтому база — −20 %.</div>")
S.append("<table><tr><th>Период</th><th>Сценарий</th><th>Заявок</th><th>Доля побед</th><th>Выручка без НДС, млн</th><th>Маржа</th><th>Валовая прибыль, млн</th></tr>")
for yn, sn, n, w, red, m, rev, gp in SC:
    S.append(f"<tr class='{'ok' if 'Базовый' in sn else ''}'><td>{yn}</td><td>{sn}</td><td class=n>{n}</td><td class=n>{w*100:.0f} %</td><td class=n>{rev/1e6:.0f}</td><td class=n>{m*100:.0f} %</td><td class=n>{gp/1e6:.0f}</td></tr>")
S.append(f"</table><div class='small'>Средний лот = средняя НМЦ лота 0,5–30 млн по 223-ФЗ за 12 мес. ({avg/1e6:.1f} млн с НДС). Доля побед 12–17 % — допущение для новичка. Маржа валовая (товар + доставка), до постоянных расходов. "
         "Выборка маржи — 6 лотов; по мере расчёта новых лотов цифру нужно обновлять. Нефтегаз на закрытых площадках в сценарий не включён (верхний потенциал). Постоянные расходы ≈12–20 млн/год → в базовом сценарии год 1 — около нуля (18 млн валовой против 12–20 млн расходов), прибыль со второго года (44 млн валовой).</div>")
# сценарий пользователя: 3 заявки в неделю, средний контракт 20–30 млн
S.append("<h2>2а. Целевой сценарий: 3 заявки в неделю, средний контракт 20–30 млн</h2>")
S.append("<table><tr><th>Средний контракт (НМЦ с НДС)</th><th>Доля побед</th><th>Побед в год</th><th>Выручка без НДС, млн</th><th>Валовая прибыль при 33 % (база), млн</th><th>при 24 % (осторожно), млн</th><th>Оборотка на 2–3 контракта параллельно, млн</th></tr>")
for lot in (20e6, 25e6, 30e6):
    for w in (.12, .15, .20):
        wins = 156 * w; rev = wins * lot * 0.8 / 1.22
        S.append(f"<tr class='{'ok' if lot == 25e6 and w == .15 else ''}'><td class=n>{lot/1e6:.0f}</td><td class=n>{w*100:.0f} %</td><td class=n>{wins:.0f}</td><td class=n>{rev/1e6:.0f}</td><td class=n>{rev*mg(.2)/1e6:.0f}</td><td class=n>{rev*mg(.3)/1e6:.0f}</td><td class=n>{2.5*lot*0.8/1.22*(1-mg(.2))/1e6:.0f}</td></tr>")
S.append("</table><div class='box'><b>Реалистичность.</b> 156 заявок в год (3 в неделю) — нормальная нагрузка для 2 тендерных специалистов. Но крупных лотов в ЕИС по нише мало: за 12 мес. в 223-ФЗ лотов 10–20 млн ≈ 74, 20–30 млн ≈ 24, 30–60 млн ≈ 42 (в основном водоподготовка и иониты); основная масса (≈870) — до 5 млн. "
         "Чтобы держать средний контракт 20–30 млн, нужны: (1) все крупные ЕИС-лоты ниши (≈140/год от 10 до 60 млн — подавать почти во все); (2) закрытые площадки нефтегаза (Роснефть, Газпром нефть, ЛУКОЙЛ, СИБУР) — там лоты реагентов и катализаторов по 20–150 млн, нужна предквалификация (2–6 мес.); "
         "(3) рамочные годовые договоры энергокомпаний (иониты, реагенты на год — 20–50 млн). Реально в год 1 — смесь: 1–2 заявки/нед. на крупные лоты + мелкие ради референсов; полный сценарий «3 в неделю по 20–30 млн» — с года 2, после допусков к нефтегазу и дилерских договоров. "
         "Оборотка: контракт 25 млн = закупка ≈11 млн без НДС; при 2–3 контрактах одновременно — 25–35 млн оборотных средств + гарантии.</div>")
S.append("<h2>3. Команда по этапам</h2><table><tr><th>Период</th><th>Штат</th></tr>"
         "<tr><td>0–3 мес.</td><td>директор/коммерческий; тендерный специалист; технолог (договор); бухгалтерия и таможенный брокер — аутсорс</td></tr>"
         "<tr><td>3–12 мес.</td><td>+ 2-й тендерный специалист; логист-ВЭД; менеджер по ключевым заказчикам (энергетика)</td></tr>"
         "<tr><td>12–24 мес.</td><td>+ штатный технолог-химик (нефтепромысловая химия, ОПИ); менеджер по нефтегазу; кладовщик</td></tr>"
         "<tr><td>24–36 мес.</td><td>+ специалист по катализаторам; инженер шеф-монтажа (или партнёр); юрист по договорам/претензиям</td></tr></table>")
S.append("<h2>4. Ключевые показатели (KPI)</h2><ul><li>Поданных заявок в месяц: 12–15 (год 1) → 25 (год 2) → 35+ (год 3).</li><li>Доля побед: ≥12 % / ≥15 % / ≥17 %.</li><li>Дилерских договоров: 8 / 15 / 25; доля российских производителей с записями ПП 719 — ≥50 %.</li>"
         "<li>Маржа при цене победы: ≥25 % (база по расчёту лотов ≈33 %); ни одного контракта в минус кроме согласованных «ради референса».</li><li>Срывов поставок и штрафов: 0 (риск РНП).</li><li>Допусков к закрытым площадкам нефтегаза: 0 / 3 / 5.</li></ul>")
S.append("<h2>5. Главные риски и ответ на них</h2><table><tr><th>Риск</th><th>Ответ</th></tr>"
         "<tr><td>Нет писем производителей — проигрыш по «полномочиям» и «без эквивалента»</td><td>Дилерские договоры заранее; таблица писем с датами; российские производители с ПП 719</td></tr>"
         "<tr><td>Демпинг конкурентов-дилеров</td><td>Прямой импорт из Китая (иониты, угли, гидразин) — себестоимость ниже российских дилеров</td></tr>"
         "<tr><td>Кассовый разрыв (отсрочки 30–90 дней, предоплата Китаю)</td><td>Кредитная линия, факторинг, гарантии; не брать больше 2–3 крупных лотов одновременно на старте</td></tr>"
         "<tr><td>Брендовые реагенты одного производителя (Оптион, Эмикор, СИНАЛАБ)</td><td>Дилерство у самих производителей или запрос разъяснений «или эквивалент»; иначе пропускать</td></tr>"
         "<tr><td>Опасные грузы, драгметаллы, санкции</td><td>Перевозчики с ДОПОГ, ДМДК до работы с катализаторами Pt/Pd, не обещать ушедшие бренды</td></tr></table>")
S.append("<div class='small'>Связанные файлы: PLAN_UPAKOVKI_KOMPANII.pdf, NISHA_CHASTOTA.pdf/.xlsx, NISHA_MARZHA.pdf/.xlsx, NISHA_KONKURSY_DILERSTVA.xlsx, ZAPROSY_DILERSTVO.docx.</div>")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/razv.html"
open(hp, "w").write(f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body>" + "".join(S) + "</body></html>")
html_to_pdf(hp, "../PLAN_RAZVITIYA_KOMPANII.pdf")
print(round(avg), [(x[0][:5], x[1][:6], round(x[5],3), round(x[6] / 1e6), round(x[7] / 1e6)) for x in SC])
