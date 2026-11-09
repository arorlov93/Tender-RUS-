"""NVIDIA B300 / B200 / H200: рынок и цены (09.11.2026) + поставка от ресейлеров США в Польшу против покупки в ЕС.
Источники — URL в таблице S. Фрахт, страховка, брокер — ОЦЕНКА. → ../GPU_B300_RYNOK_POLSHA_0911.xlsx + .pdf"""
import html, openpyxl
from openpyxl.styles import Font, PatternFill
from pdfutil import html_to_pdf
from tidy import tidy

USD_RUB, EUR_RUB = 82.1386, 92.1924          # ЦБ РФ 07.11.2026
EURUSD = EUR_RUB / USD_RUB                    # ≈ 1,122
PLN_EUR = 3023402 / 709787                    # курс из цен Servermall PL/EU ≈ 4,26

# Логистика США → Польша, один 8U-сервер HGX (ОЦЕНКА): брутто в ящике ~230 кг, авиа $5–7/кг с топливом и охраной ценного груза
def logistics_usd(value_usd):
    air = 230 * 6.0 + 600            # авиа + обработка ценного груза (оценка)
    ins = value_usd * 0.004          # страховка 0,3–0,5 % (оценка)
    us_exp = 350                     # экспортное оформление США, подача EEI/AES (оценка)
    pl_broker = 450 * EURUSD         # брокер в Польше, ЕАД, СВХ (оценка)
    truck = 350 * EURUSD             # аэропорт Варшавы → ЦОД (оценка)
    return air + ins + us_exp + pl_broker + truck

# (модель, где, цена, валюта, без/с НДС, статус, источник)
S = [
 ("Supermicro SYS-822GS-NB3RT, HGX B300 (база)", "США, Supermicro eStore", 536771, "USD", "без налога", "под заказ", "store.supermicro.com/us_en/8u-gpu-superserver-sys-822gs-nb3rt.html"),
 ("Supermicro SYS-822GS-NB3RT, кастом", "США, AVADirect", 535064, "USD", "без налога", "под заказ", "avadirect.com (конфигуратор)"),
 ("Supermicro HGX B300, фикс. конфигурация", "США, сводка gpuperhour", 663000, "USD", "без налога", "—", "gpuperhour.com/blog/nvidia-gpu-prices"),
 ("Supermicro SYS-822GS-NB3RT (полная)", "США, Servermall", 794962, "USD", "без налога", "—", "servermall.com/en-us/catalog/servers/supermicro-sys-822gs-nb3rt-hgx-b300-nvl8-8nvme-1/"),
 ("Lenovo SR680a V4, 8×B300", "США, Lenovo", 785606, "USD", "без налога", "—", "intuitionlabs.ai / thundercompute.com"),
 ("Supermicro SYS-822GS-NB3RT HGX B300", "ЕС, Servermall (доставка по ЕС включена)", 709787, "EUR", "без НДС", "в наличии (по сайту)", "servermall.com/catalog/servers/supermicro-sys-822gs-nb3rt-hgx-b300-nvl8-8nvme/"),
 ("То же, польская витрина", "PL, Servermall", 3023402, "PLN", "без НДС", "в наличии (по сайту)", "servermall.com/pl/catalog/servers/supermicro-sys-822gs-nb3rt-hgx-b300-nvl8-8nvme/"),
 ("Supermicro AS-8126GS-NB3RT HGX B300 (AMD)", "PL, Million Miner", None, "—", "—", "цена по запросу", "millionminer.com/pl/product/ai-hardware/supermicro-gpu-a-server-as-8126gs-nb3rt-nvidia-hgx-b300-nvl8"),
]
R = []
for n, where, p, cur, vat, st, src in S:
    if p is None:
        R.append(dict(n=n, w=where, p="по запросу", st=st, src=src, net=None, gross=None, eu=True)); continue
    usd = p if cur == "USD" else p * EURUSD if cur == "EUR" else p / PLN_EUR * EURUSD
    eu = cur != "USD"
    net_eur = (usd / EURUSD) if eu else (usd + logistics_usd(usd)) / EURUSD
    R.append(dict(n=n, w=where, p=f"{p:,.0f} {cur}".replace(",", " "), st=st, src=src, net=net_eur, gross=net_eur * 1.23, eu=eu, log=None if eu else logistics_usd(usd) / EURUSD))

MKT = [  # рынок 8-GPU серверов, США
 ("H200 HGX, 8 GPU", "$319–398 тыс.", "≈ $40–50 тыс.", "в наличии / недели", "gpuperhour.com, enzu.com"),
 ("B200 HGX, 8 GPU", "$384–550 тыс.", "≈ $48–69 тыс.", "BIZON: отгрузка с середины января", "gpuperhour.com, nodiligence.substack.com"),
 ("B300 HGX, 8 GPU", "$537–663 тыс. (реселлеры до $760–830 тыс.)", "≈ $50–65 тыс.", "8–20 недель; UpStation: 30+ недель", "gpuperhour.com, enzu.com, gpusmith.com, nodiligence"),
 ("B300 в Китае (серый рынок, Reuters)", "≈ $1 млн", "—", "—", "thenextweb.com (Reuters)"),
]

wb = openpyxl.Workbook(); ws = wb.active; ws.title = "B300 в Польшу"
ws.append(["Модель", "Где", "Цена продавца", "Наличие", "Логистика США→PL, € (оценка)", "Итого в Польше без НДС, €", "С НДС 23 %, €", "Источник"])
for r in R:
    ws.append([r["n"], r["w"], r["p"], r["st"], round(r["log"]) if r.get("log") else ("включена" if r["eu"] and r["net"] else "—"),
               round(r["net"]) if r["net"] else "—", round(r["gross"]) if r["gross"] else "—", r["src"]])
s2 = wb.create_sheet("Рынок США"); s2.append(["Сервер", "Цена сервера", "≈ за 1 GPU", "Срок поставки", "Источник"])
for m in MKT: s2.append(list(m))
s3 = wb.create_sheet("Допущения"); s3.append(["Статья", "Значение", "Комментарий"])
for row in [("Курс", f"€1 = ${EURUSD:.3f}; €1 = {PLN_EUR:.2f} PLN", "ЦБ РФ 07.11.2026; PLN — из цен Servermall"),
            ("Вес сервера в ящике", "≈ 230 кг брутто", "8U, оценка"),
            ("Авиафрахт США → Варшава", "$6/кг + $600 обработка ценного груза", "рынок $2,5–5,5/кг + надбавки, оценка"),
            ("Страховка", "0,4 % стоимости", "оценка"),
            ("Экспортное оформление США (EEI/AES)", "$350", "оценка"),
            ("Брокер PL + доставка до ЦОД", "€450 + €350", "оценка"),
            ("Пошлина ЕС, код 8471 (ITA)", "0 %", "celna24.pl/customs-tariff/8471"),
            ("НДС импорта Польши", "23 %", "возмещается плательщику НДС; можно art. 33a (без уплаты на таможне)"),
            ("Экспортный контроль США", "ECCN 4A090.a (серверы) / 3A090.a (GPU)", "нужна проверка лицензии; гайданс BIS 31.05.2026"),
            ("Санкции ЕС, ст. 12g Рег. 833/2014", "оговорка «no re-export to Russia» в договоре", "серверы 8471 — товары высокого приоритета")]:
    s3.append(list(row))
for s in wb:
    for c in s[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")
for i, w in enumerate([40, 30, 18, 18, 16, 16, 14, 50], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
for i, w in enumerate([30, 36, 16, 30, 40], 1): s2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
for i, w in enumerate([34, 40, 50], 1): s3.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
tidy(wb, row_h=30); wb.save("../GPU_B300_RYNOK_POLSHA_0911.xlsx")

e = html.escape
css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:8.5px;color:#111;margin:0}h1{font-size:13px;margin:0 0 4px}h2{font-size:10px;margin:6px 0 2px}
table{border-collapse:collapse;width:100%}th,td{border:1px solid #bbb;padding:2px 4px;vertical-align:top}th{background:#305496;color:#fff;text-align:left}
td.n{text-align:right;white-space:nowrap}.box{border:1px solid #305496;padding:4px 6px;background:#f3f6fb;margin:3px 0}.s{font-size:7.5px;color:#444}.ok{background:#c6efce}.warn{background:#fff2cc}"""
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body>",
     "<h1>NVIDIA B300: рынок, цены и поставка в Польшу из США и из ЕС (09.11.2026)</h1>",
     "<div class='box'><b>Вывод.</b> Сервер HGX B300 (8 GPU) в США стоит $537–663 тыс., в ЕС у Servermall €710 тыс. без НДС (в наличии по сайту). "
     "С доставкой из США в Польшу (оценка ~€5–6 тыс. на сервер, ~1 %) сервер выходит <b>≈ €483–596 тыс. без НДС</b>, т.е. на €115–225 тыс. дешевле ЕС, "
     "но срок 8–30 недель и нужна проверка экспортной лицензии США. Пошлина ЕС 0 %, НДС 23 % (возмещается). "
     "«Отдельных карт» B300 нет: это модули SXM на плате HGX, продаются в составе сервера; если поляки хотят карты PCIe, это RTX PRO 6000 Blackwell или H200 NVL.</div>",
     "<h2>1. Рынок 8-GPU серверов (США)</h2><table><tr><th>Сервер</th><th>Цена</th><th>≈ за GPU</th><th>Срок</th><th>Источник</th></tr>"]
for m in MKT: h.append("<tr>" + "".join(f"<td>{e(x)}</td>" for x in m) + "</tr>")
h.append("</table><h2>2. Сервер HGX B300 в Польше: купить в США и привезти или купить в ЕС</h2><table><tr><th>Модель</th><th>Где</th><th>Цена продавца</th><th>Наличие</th>"
         "<th>Логистика США→PL (оценка)</th><th>В Польше без НДС</th><th>С НДС 23 %</th><th>Источник</th></tr>")
for r in R:
    lg = f"€{r['log']:,.0f}".replace(",", " ") if r.get("log") else ("включена" if r["net"] else "—")
    f = lambda v: f"€{v:,.0f}".replace(",", " ") if v else "—"
    h.append(f"<tr><td>{e(r['n'])}</td><td>{e(r['w'])}</td><td class=n>{e(r['p'])}</td><td>{e(r['st'])}</td><td class=n>{lg}</td>"
             f"<td class='n {'warn' if r['eu'] else 'ok'}'><b>{f(r['net'])}</b></td><td class=n>{f(r['gross'])}</td><td class=s>{e(r['src'])}</td></tr>")
h.append("</table><h2>3. Как везти из США в Польшу</h2><div class='box'>"
         "<b>Маршрут:</b> ресейлер (США) → авиа до Варшавы (WAW) 3–7 дней, ящик ~230 кг на сервер → таможня PL → ЦОД. Условия: FCA/EXW США или DAP Польша. "
         "<b>Затраты на сервер (оценка):</b> авиа ~$2 000, страховка 0,4 % (~$2 200), экспорт США $350, брокер PL €450, грузовик €350. "
         "<b>Пошлина</b> 0 % (8471, ITA); <b>НДС</b> 23 % от CIF, вычитается польским плательщиком НДС (art. 33a — без уплаты на границе).<br>"
         "<b>Экспортный контроль США:</b> серверы B300 — ECCN 4A090.a. Польша в группе A:5, но правила по чипам ИИ в 2025–2026 менялись; по гайдансу BIS от 31.05.2026 лицензия по-прежнему нужна в ряде случаев "
         "(в т.ч. если у покупателя материнская компания в группе D:5). Ресейлер США подаст EEI и запросит у поляков end-user statement: кто конечный пользователь, адрес ЦОД, назначение. "
         "Без реального конечного пользователя в Польше сделки не будет. <b>ЕС:</b> в договоре обязательна оговорка о запрете реэкспорта в Россию (ст. 12g Рег. 833/2014).<br>"
         "<b>Что запросить у польского клиента:</b> название компании, NIP и KRS, адрес ЦОД, назначение, модель (сервер HGX или PCIe-карты), количество, срок, готовность подписать end-user statement.</div>"
         "<p class=s>Цены — открытые витрины и обзоры на ноябрь 2026, официального прайса NVIDIA нет; финальная цена только по КП. Курс ЦБ РФ 07.11.2026: €1 = $1,122.</p></body></html>")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/gpu.html"
open(hp, "w").write("".join(h)); html_to_pdf(hp, "../GPU_B300_RYNOK_POLSHA_0911.pdf")
for r in R: print(r["n"][:40], r["p"], round(r["net"] or 0), round(r.get("log") or 0))
