"""Итоговое решение по закупкам РФ с присланными ТЗ (16.11.2026) → ../RF_RESHENIE_1611.xlsx + .pdf"""
import json, os, html, subprocess
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as CL
from isp_letterhead import SCRATCH
Z = "data/zakupki_1511_tz/"
def eco(f):
    e = json.load(open(Z + f + "/analysis.json")).get("economics") or {}
    return e.get("profit_nmc"), e.get("profit_m10"), e.get("profit_m20")
# (папка, номер, заказчик/предмет, подача, НМЦ ₽, решение, условие/причина, кому запросы КП)
GO = [
 ("kuzbass_32616549223", "32616549223", "Кузбассэнерго-РЭС — видеорегистраторы Страж-ПВР-05, 17 шт", "18.11 10:00", 629125, "Подаём", "КП производителя «Спектр» (Омск)", "Спектр"),
 ("sakh_0361200015026004877", "0361200015026004877", "ЮСГКБ (Сахалин), 44-ФЗ — пробирки вакуумные 4 мл, 46 150 шт", "18.11 01:00 МСК", 960843, "Подаём при условии", "пробирка с записью в реестре ПП 719 + РУ; доставка до Южно-Сахалинска", "Еламед, Беримед, Гранат/ОМБ, Юнимед/Церебрум, ЗМТ"),
 ("hakas_32616563695", "32616563695", "Хакасэнерго — ОПС/видеонаблюдение, 7 позиций", "19.11 11:00", 734215, "Подаём при условии", "подтвердить цены дистрибьютора DSSL/TRASSIR", "DSSL (письмо из разбора 15.11)"),
 ("lipetsk_32616572997", "32616572997", "Липецк — катионит КУ-2-8 Na в/с, 16 т", "20.11 09:00", 3445583, "Подаём при условии", "ТОКЕМ (единственный с записью ПП 719) даёт ≤ ~170 ₽/кг; иначе убыток", "ТД ТОКЕМ, Аквахим, Химэнерго"),
 ("stupino_32616577851", "32616577851", "ПТО ЖКХ Ступино — сульфат алюминия 1 сорт, 260 т", "23.11 14:00", 15860000, "Подаём при условии", "сульфат ≤ ~49 тыс. ₽/т фикс до 11.2027; гарантия 793 тыс.; регистрация ПИК ЕАСУЗ; обеспечение заявки 158,6 тыс.", "Аурат, Аквахим, Сорбент, Урал-Коагулянт, Аква-Реагент, Химпродукт"),
 ("energoneft_32616571048", "32616571048", "Энергонефть Томск — химматериалы 38 поз., Стрежевой", "23.11", 5292098, "Подаём при условии", "КП МедЭкоТест (тест-системы МЭТ — основная маржа) и галит; скачать извещение (сравнение с/без НДС, оплата)", "МедЭкоТест, Солена, Тюменская Соляная, Аквахим, Реактив"),
 ("ugra_32616559738", "32616559738", "Югорская ЭК — стройматериалы 92 поз., Ханты-Мансийск", "23.11 06:00", 1835717, "Подаём при условии", "КП по 10 крупнейшим позициям + номера ПП 719; половина цен — оценка", "ТСК-групп, Строительный Двор, Магистраль НефтеГаз, Benza, Сервис-Петролиум"),
 ("dgk_32616569450", "32616569450", "ДГК — флокулянт Праестол 4 075 кг, Дальний Восток", "23.11 10:00", 3783097, "Подаём при условии", "оригинальный Праестол ≤ ~720 ₽/кг с НДС + номера ГИСП; аккредитация РусГидро", "Аквахим, Акрил-Химпроект (из разбора)"),
 ("novgor_0550600003626000003", "0550600003626000003", "Новгородский р-н, 44-ФЗ — сульфат алюминия 26,4 т + гипохлорит 3,78 т", "23.11 09:20", 1995161, "Подаём при условии", "маржа тонкая; поставка до 31.12; не ниже ~1,75 млн", "— (запросы не отправлялись)"),
 ("viza_32616569486", "32616569486", "ТЭЦ ВИЗа (Екатеринбург) — ПАА-гель 10,8 т", "24.11 04:00", 1522080, "Подаём при условии", "допустят ли гель других ТУ; нужны извещение и ТЗ", "СИНТЕК (звонок), НеваРеактив, Химальянс, Аквахим, АМК-Групп"),
]
NO = [
 ("32616539091", "Водоканал Казань — ПОХА 2 100 т", "19.11", 128100000, "маржа ≈0 по НМЦ, оплата 90 дн., штраф 50 %"),
 ("32616567340", "Водоочистка Архангельск — жидкий сульфат Al", "24.11", 2950000, "убыток уже по НМЦ (торги без НДС)"),
 ("32616575628", "Водоканал Киров — Гринлайф К10 ПВ 13 т", "24.11", 4769007, "допуск только после испытаний; рынок 792 ₽/кг при НМЦ 367"),
 ("32616571274", "Ульяновскводоканал — Гринлайф К10 ПВ", "30.11", 2762850, "эквивалент запрещён; убыток 3,8 млн"),
 ("32616571917", "Ямалкоммунэнерго — 7 реагентов", "30.11", 9202141, "бренды зафиксированы, нужен дилерский статус; убыток"),
 ("32616567369", "Леноблводоканал — ПОХА 379 т", "27.11", 37207657, "10,5 % по НМЦ; с НДС 5 % при −10 % убыток; договор после тарифов"),
 ("32616573584", "Кировэнерго — расходники СДТУ 73 поз.", "20.11", 4570584, "прибыль 62 тыс. при −10 %, много рисков"),
 ("32616554083", "Россети Томск — каталог 526 поз.", "19.11", 1874000, "убыток при −10 %"),
 ("32616557953", "Интер РАО Приморская ТЭС — мембранная химия СИНАЛАБ", "23.11", 1837893, "убыток при −10 %; эквивалент — СГР + 3 договора"),
 ("32616548382", "Интер РАО Прегольская ТЭС", "13.11", 1314138, "срок подачи истёк"),
 ("32616569002", "Башкирэнерго — телеком-шкафы", "20.11", 1217900, "убыток по НМЦ (сравнение без НДС)"),
 ("0321500002026000241", "Стававтодор — запчасти, 44-ФЗ", "20.11", 1600000, "нет каталожных номеров, срок 2 недели"),
 ("32616545637", "Вторресурсхолдинг — канцтовары", "23.11", 1000000, "цены ТЗ = розница Комуса, убыток при −10 %"),
 ("0862300028026000019", "Реж — 7 реагентов, 44-ФЗ", "?", 2064600, "актуальное извещение не найдено; убыток при −10 %"),
]
WAIT = [("32616578855", "Россети ПМЭС — химпродукция 2027", "23.11", 15178745, "документации нет (присланный архив — Энергонефть); скачать из ЕИС")]
link = lambda n: (f"https://zakupki.gov.ru/epz/order/notice/notice223/common-info.html?regNumber={n}" if len(n) == 11 else
                  f"https://zakupki.gov.ru/epz/order/extendedsearch/results.html?searchString={n}")
H = PatternFill("solid", fgColor="0B2545"); HF = Font(bold=True, color="FFFFFF"); W = Alignment(wrap_text=True, vertical="top"); LINK = Font(color="0563C1", underline="single")
G = PatternFill("solid", fgColor="E6F4EA"); Y = PatternFill("solid", fgColor="FFF4D6"); R = PatternFill("solid", fgColor="FCE8E8")
wb = Workbook(); ws = wb.active; ws.title = "Подаём"
ws.append(["№", "Номер (ссылка)", "Заказчик / предмет", "Подача", "НМЦ, ₽", "Прибыль при НМЦ", "Прибыль −10 %", "Прибыль −20 %", "Решение", "Условие участия", "Запросы КП направлены"])
for c in ws[1]: c.fill = H; c.font = HF; c.alignment = W
tot = 0
for i, (f, n, s, dl, nmc, v, cond, sup) in enumerate(GO, 1):
    p0, p10, p20 = eco(f); tot += p10 if isinstance(p10, (int, float)) and p10 > 0 else 0
    ws.append([i, n, s, dl, nmc, p0, p10, p20, v, cond, sup])
    c = ws.cell(ws.max_row, 2); c.hyperlink = link(n); c.font = LINK
    for cc in (5, 6, 7, 8): ws.cell(ws.max_row, cc).number_format = "# ##0"
    for cc in range(1, 12): ws.cell(ws.max_row, cc).fill = G if v == "Подаём" else Y
ws.append([]); ws.append(["", "", "Потенциал прибыли при −10 % (если все условия выполнены)", "", "", "", round(tot)]); ws.cell(ws.max_row, 3).font = Font(bold=True); ws.cell(ws.max_row, 7).number_format = "# ##0"
for i, w in enumerate([4, 22, 46, 15, 14, 13, 13, 13, 16, 56, 42], 1): ws.column_dimensions[CL(i)].width = w
for row in ws.iter_rows(min_row=2):
    for c in row: c.alignment = W
ws.freeze_panes = "C2"
s2 = wb.create_sheet("Не подаём")
s2.append(["Номер (ссылка)", "Заказчик / предмет", "Подача", "НМЦ, ₽", "Причина"])
for c in s2[1]: c.fill = H; c.font = HF; c.alignment = W
for n, s, dl, nmc, why in NO + WAIT:
    s2.append([n, s, dl, nmc, why]); c = s2.cell(s2.max_row, 1); c.hyperlink = link(n); c.font = LINK; s2.cell(s2.max_row, 4).number_format = "# ##0"
    for cc in range(1, 6): s2.cell(s2.max_row, cc).fill = R if (n, s, dl, nmc, why) in NO else Y
for i, w in enumerate([22, 50, 10, 14, 70], 1): s2.column_dimensions[CL(i)].width = w
for row in s2.iter_rows(min_row=2):
    for c in row: c.alignment = W
s3 = wb.create_sheet("Как считано")
for x in ["Прибыль — после УСН 15 % (доходы минус расходы, минимум 1 % от дохода), без НДС; себестоимость — открытые прайсы/публикации + доставка (оценка). Не КП.",
          "«Подаём при условии» — экономика держится только при подтверждении цены/документов поставщика (графа «Условие участия»).",
          "Детали по позициям: RF_MARZHA_TZ_1611.xlsx; тексты запросов и спецификации: ZAPROSY_KP_1611.zip."]: s3.append([x])
s3.column_dimensions["A"].width = 150
wb.save("../RF_RESHENIE_1611.xlsx")
# PDF
e = html.escape; k = lambda v: "—" if not isinstance(v, (int, float)) else f"{v:,.0f}".replace(",", " ")
css = ("@page{size:A4 landscape;margin:7mm}body{font-family:'DejaVu Sans',Arial;font-size:7.3px;color:#111}h1{font-size:13px;margin:0 0 3px;color:#0B2545}h2{font-size:9.5px;margin:6px 0 2px;color:#0B2545}"
       "table{border-collapse:collapse;width:100%}th{background:#0B2545;color:#fff;padding:2px 3px;text-align:left}td{border-bottom:1px solid #ddd;padding:2px 3px;vertical-align:top}.n{text-align:right;white-space:nowrap}"
       ".g{background:#E6F4EA}.y{background:#FFF4D6}.r{background:#FCE8E8}a{color:#0563C1;text-decoration:none}.s{color:#555}")
h = [f"<!DOCTYPE html><html><head><meta charset=utf-8><style>{css}</style></head><body><h1>Куда подаёмся — закупки РФ по присланным ТЗ (ООО «Гидропас», 16.11.2026)</h1>",
     "<div class=s>Прибыль после УСН 15 %, без НДС; себестоимость — открытые прайсы + доставка (оценка), не КП. Номер — ссылка на ЕИС.</div><h2>Подаём (10)</h2><table><tr><th>№</th><th>Номер</th><th>Заказчик / предмет</th><th>Подача</th><th>НМЦ</th><th>При НМЦ</th><th>−10 %</th><th>−20 %</th><th>Решение</th><th>Условие участия</th><th>Запросы КП</th></tr>"]
for i, (f, n, s, dl, nmc, v, cond, sup) in enumerate(GO, 1):
    p0, p10, p20 = eco(f)
    h.append(f"<tr class={'g' if v == 'Подаём' else 'y'}><td>{i}</td><td><a href='{link(n)}'>{n}</a></td><td>{e(s)}</td><td>{e(dl)}</td><td class=n>{k(nmc)}</td><td class=n>{k(p0)}</td><td class=n><b>{k(p10)}</b></td><td class=n>{k(p20)}</td><td>{e(v)}</td><td>{e(cond)}</td><td class=s>{e(sup)}</td></tr>")
h.append(f"</table><div><b>Потенциал прибыли при −10 %: {k(tot)} ₽</b> (если все условия выполнены).</div><h2>Не подаём (14) и отложено (1)</h2><table><tr><th>Номер</th><th>Заказчик / предмет</th><th>Подача</th><th>НМЦ</th><th>Причина</th></tr>")
for n, s, dl, nmc, why in NO:
    h.append(f"<tr class=r><td><a href='{link(n)}'>{n}</a></td><td>{e(s)}</td><td>{e(dl)}</td><td class=n>{k(nmc)}</td><td>{e(why)}</td></tr>")
for n, s, dl, nmc, why in WAIT:
    h.append(f"<tr class=y><td><a href='{link(n)}'>{n}</a></td><td>{e(s)}</td><td>{e(dl)}</td><td class=n>{k(nmc)}</td><td>{e(why)}</td></tr>")
h.append("</table></body></html>")
hp = f"{SCRATCH}/rf_reshenie.html"; open(hp, "w").write("".join(h))
subprocess.run(["/opt/pw-browsers/chromium", "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer", f"--print-to-pdf={os.path.abspath('../RF_RESHENIE_1611.pdf')}", "file://" + hp], check=True, capture_output=True, timeout=180)
import pypdf; print("pages", len(pypdf.PdfReader("../RF_RESHENIE_1611.pdf").pages), "tot", tot)
