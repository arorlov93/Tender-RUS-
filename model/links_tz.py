"""Таблица 13 лотов: ссылки на закупки + где нужно достать ТЗ и что в нём проверить. → ../SSYLKI_TZ_0611.xlsx"""
import json, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from tidy import tidy
S = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/live/"
DT = json.load(open(S + "det.json"))
it = {r[0]: r for r in openpyxl.load_workbook("../ITOG_CENY_MARZHA_0611.xlsx").active.iter_rows(min_row=2, values_only=True)}
kp = {r[1]: r for r in openpyxl.load_workbook("../KONKURSY_KP_0611.xlsx")["Конкурсы (13)"].iter_rows(min_row=2, values_only=True)}
B2B = {"32616536765": "4665013", "32616541198": "4666969", "32616542934": "4667633", "32616548235": "4669950"}
# приоритет и что смотреть в ТЗ
TZ = {
 "32616523325": ("1 — главный", "ДА, срочно", "Почему НМЦ 476 ₽/кг (марка, концентрация, требования к теплоносителю, упаковка), объём партий и график, сравнение с НДС или без, требования к опыту/аттестации"),
 "32616534213": ("1", "ДА, срочно", "Марка ИОМС-1 обязательна или «эквивалент»; «в пересчёте на 100 %»; график поставок 2027; нацрежим «ограничение» — нужна запись ПП 719"),
 "32616548235": ("1", "ДА", "Точное количество АН-31 (в ЕИС нет), ГОСТ/ТУ, график поставки, допускается ли эквивалент"),
 "32616534438": ("2", "ДА", "Эмикор без эквивалента или с эквивалентом; что делать, если бренд «Эмикор» только у одного держателя ТУ"),
 "32616522434": ("2", "ДА", "Марка аммиака (А/Б), график 700 т (разово или помесячно), тара/цистерны, обеспечение заявки 105,6 тыс. и исполнения 1,06 млн, оплата"),
 "32616539091": ("2", "ДА", "Марка ПОХА (Аква-Аурат/иное, % Al2O3), график 2 100 т, ж/д или авто, отсрочка оплаты, обеспечение 1,28 + 6,4 млн"),
 "32616531631": ("2", "ДА, обязательно", "Марка Silcarbon (S814/S835 — идём; K814 — нет), эквивалент допускается?"),
 "32616548412": ("3", "желательно", "Вид соли (таблетированная, ГОСТ), фасовка, адрес, сроки — лот быстрый, прибыль небольшая"),
 "32616541285": ("3", "желательно", "Сравнение с НДС или без (решает маржу на УСН), чистота HEDP 98 %, тара"),
 "32616541198": ("3", "желательно", "Только если КП Сорбента ≤190 ₽/кг; график поставки"),
 "32616536765": ("4", "ДА, если идти", "Марка и количество «ИОС Катионит» (в ЕИС нет) — без этого не посчитать"),
 "32616549489": ("пропуск", "нет", "Маржа ≈4 % при −10 % — не идём"),
 "32616542934": ("пропуск", "нет", "Убыток — не идём"),
}
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Ссылки и ТЗ"
H = ["Приоритет", "Нужно достать ТЗ", "Что проверить в ТЗ", "Подача до", "№ ЕИС", "Заказчик", "Товар", "НМЦ, ₽", "Маржа УСН при −10 %", "Площадка", "Ссылка ЕИС (карточка + документы)", "Ссылка на площадке", "Файлы документации"]
ws.append(H)
order = sorted(it, key=lambda n: (TZ.get(n, ("9",))[0][:1] if TZ.get(n, ("9",))[0][0].isdigit() else "9", str(kp.get(n, [""])[0])))
for n in order:
    p, need, what = TZ.get(n, ("", "", ""))
    k = kp.get(n); r = it[n]; d = DT.get(n, {})
    eis = f"https://zakupki.gov.ru/epz/order/notice/notice223/common-info.html?regNumber={n}"
    docs = f"https://zakupki.gov.ru/epz/order/notice/notice223/documents.html?regNumber={n}"
    etp_url = f"https://www.b2b-center.ru/market/view.html?id={B2B[n]}" if n in B2B else (d.get("url") or "")
    ws.append([p, need, what, (str(k[0])[:16] if k else ""), n, r[2], r[3], r[6], r[12], r[23] or (d.get("etp") or ""), docs, etp_url, k[17] if k else ""])
    X = ws.max_row
    ws.cell(X, 8).number_format = "#,##0"; ws.cell(X, 9).number_format = "0%"
    ws.cell(X, 11).hyperlink = docs
    if etp_url.startswith("http"): ws.cell(X, 12).hyperlink = etp_url
    col = "f8cbad" if need.startswith("ДА") else "fff2cc" if need == "желательно" else "d9d9d9"
    for j in (1, 2, 3): ws.cell(X, j).fill = PatternFill("solid", fgColor=col)
for c in ws[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
for i, w in enumerate([12, 15, 60, 16, 13, 22, 40, 13, 10, 26, 50, 45, 45], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = "D2"; ws.auto_filter.ref = f"A1:M{ws.max_row}"
s2 = wb.create_sheet("Как скачать ТЗ")
for line in ["1. Открыть «Ссылка ЕИС» → вкладка «Документы» → скачать архив/файлы (без регистрации, доступно всем).",
             "2. Если в ЕИС файлов нет или они не открываются — «Ссылка на площадке» (B2B-Center, ТЭК-Торг, ЭТП ГПБ, РТС): для скачивания часто нужна регистрация/аккредитация.",
             "3. Прислать мне архивы (zip/rar/doc) прямо в чат — я распакую, проверю требования, количества, сроки и сравнение с НДС, пересчитаю маржу.",
             "4. Красным отмечены лоты, где без ТЗ решение не принять; жёлтым — желательно; серым — не идём."]:
    s2.append([line])
s2.column_dimensions["A"].width = 140
tidy(wb, row_h=36); wb.save("../SSYLKI_TZ_0611.xlsx"); print("ok")
