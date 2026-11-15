"""Итоговая сводка 15.11.2026: РФ (разбор ТЗ), мир (маржа), отправленные письма, партнёры, план → ../ITOG_1511.pdf + .xlsx"""
import json, os, html, subprocess
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as CL
from isp_letterhead import SCRATCH
C = "data/companies/"; T = "data/tenders_1511/"; Z = "data/zakupki_1511_tz/"
S = json.load(open(C + "rfq_sent_1511.json"))
M = json.load(open(T + "mir_marzha.json"))["rows"]
RF = [("32616549223", "Кузбассэнерго", "Страж-ПВР-05, 17 шт", "18.11 10:00", "kuzbass_32616549223", "Подавать", "КП и выписка ГИСП от «Спектр» (Омск) 16.11, подача 17.11 вечером, на переторжке не ниже −18%"),
      ("0361200015026004877", "ЮСГКБ, Сахалин (44-ФЗ)", "Пробирки вакуумные 4 мл, 46 150 шт", "18.11 01:00 МСК", "sakh_0361200015026004877", "При условиях", "Еламед/Беримед: номер реестровой записи на артикул + ставка доставки ≤200 тыс. ₽ до 17.11"),
      ("32616563695", "Хакасэнерго", "ОПС/видео, 7 позиций", "19.11 11:00", "hakas_32616563695", "При условиях", "Цены DSSL: MiniNVR ≤64 тыс. ₽, камеры ≤17 тыс. ₽; прочесть 3 разъяснения"),
      ("32616573584", "Кировэнерго", "Расходники СДТУ, 73 позиции", "20.11 11:00", "kirov_32616573584", "Скорее нет", "Запрос разъяснений до 17.11 11:00 (реестр ПП 719 для импорта); Меркурий ≤15 тыс. ₽"),
      ("32616554083", "Россети Томск", "Каталог 526 позиций, без количеств", "19.11 11:00", "tomsk_32616554083", "Не подавать", "Рынок ≥ НМЦ по российским позициям, проигрыш ОСН-конкурентам")]
rf = []
for n, cust, what, dl, d, v, todo in RF:
    e = json.load(open(Z + d + "/analysis.json"))["economics"]
    rf.append(dict(n=n, cust=cust, what=what, dl=dl, v=v, todo=todo, nmc=e.get("revenue_nmc"), p10=e.get("profit_m10"), p20=e.get("profit_m20"), m10=e.get("margin_pct_m10")))
ORD = {"Напрямую (иностранцы допускаются)": 0}
top = sorted([o for o in M if o["margin_usd"] and o["margin_usd"] > 0 and o["access"].startswith("Напрямую")], key=lambda o: -o["margin_usd"])[:10]
mail = []
for k, lab in [("sent", "КП заводам — Африка"), ("batch2_eu", "КП заводам — ЕС"), ("batch3_site_emails_user_ok", "КП заводам — пакет 3"), ("batch4_africa_partners", "Партнёры Африка"), ("batch5_local_mena", "Местные участники MENA/Турция"),
               ("batch6_local_cm_eu", "Местные участники Камерун/ЕС"), ("batch7_local_africa_eu", "Местные участники Африка/ЕС")]:
    for x in S.get(k, []): mail.append(dict(group=lab, to=x["to"], what=x.get("product") or x.get("company") or "", tender=x.get("tender") or ""))
PLAN = [("16.11", "Гидропас", "Кузбасс: звонок/КП «Спектр», выписка ГИСП; Сахалин: КП Еламед/Беримед + ставка доставки; Хакасия: КП DSSL"),
        ("16.11", "ISP", "Отправить очередь 25 писем: Индия, Филиппины, Бразилия, Вьетнам, Таиланд, остаток MENA"),
        ("16.11", "ISP (решение)", "«Отправляй» по черновикам: Kairui (T-47), Rimpro, экспедитор, Suqing"),
        ("17.11", "Гидропас", "Подать Кузбасс и Сахалин (до 01:00 МСК 18.11); Киров — запрос разъяснений до 11:00"),
        ("18–19.11", "ISP", "Ответы по Кипру (FeCl3) и EMAYA (PAM, образцы до 24–25.11) → решение о подаче"),
        ("19.11", "Гидропас", "Хакасия — подача, если цены DSSL прошли"),
        ("20–21.11", "ISP", "Ответы по Уганде (сульфат, сода) и Кении (PAC); решение по агенту в Кампале и гарантии $53,5 тыс."),
        ("23.11", "ISP", "Подача: Кипр, Aquanet (если есть цена и партнёр)"), ("25–27.11", "ISP", "Подача: EMAYA (25.11), NWSC Уганда (27.11)"),
        ("01–16.12", "ISP", "HERA, Мюнстер, Укрнафта, CIIP, CAMWATER (через партнёра, 07.12), Uz-Kor T-47 (02.12 — только с полным пакетом)"),
        ("без срока", "ISP (решение)", "JSRS для Омана ($1 000/год); регистрация нерезидента zakup.sk.kz; лаб. испытания деэмульгатора Rimpro в КМГ Инжиниринг")]
e = html.escape; k = lambda v: "—" if not isinstance(v, (int, float)) else f"{v:,.0f}".replace(",", " ")
css = ("@page{size:A4 landscape;margin:8mm}body{font-family:'DejaVu Sans',Arial;font-size:7.4px;color:#111}h1{font-size:13.5px;margin:0 0 3px;color:#0B2545}h2{font-size:9.5px;margin:6px 0 2px;color:#0B2545}"
       "table{border-collapse:collapse;width:100%}th{background:#0B2545;color:#fff;padding:2px 3px;text-align:left}td{border-bottom:1px solid #ddd;padding:1px 3px;vertical-align:top}.n{text-align:right;white-space:nowrap}"
       ".g{color:#555}.go{color:#1B7F3B;font-weight:bold}.c{color:#B07D00;font-weight:bold}.no{color:#B00020;font-weight:bold}.cols{display:flex;gap:8px}.cols>div{flex:1}")
h = [f"<!DOCTYPE html><html><head><meta charset=utf-8><style>{css}</style></head><body><h1>Итог на 15.11.2026: Россия, мир, письма, план</h1>",
     f"<div class=g>Сегодня с info@ispgroupgc.com отправлено {S.get('daily_count_1511')} писем (лимит 75): 21 запрос КП заводам + 54 предложения партнёрам/местным участникам. Разобрано 5 закупок РФ по ТЗ и 163 зарубежных тендера.</div>",
     "<h2>1. Россия — разбор ТЗ (Гидропас, УСН, прибыль после налога, без НДС)</h2><table><tr><th>Номер</th><th>Заказчик / предмет</th><th>Подача</th><th>НМЦ, ₽</th><th>Прибыль −10%</th><th>−20%</th><th>Решение</th><th>Что сделать</th></tr>"]
for r in rf:
    cls = "go" if r["v"] == "Подавать" else "no" if r["v"] == "Не подавать" else "c"
    h.append(f"<tr><td>{r['n']}</td><td><b>{e(r['cust'])}</b> — {e(r['what'])}</td><td>{e(r['dl'])}</td><td class=n>{k(r['nmc'])}</td><td class=n>{k(r['p10'])} ({r['m10']}%)</td><td class=n>{k(r['p20'])}</td><td class={cls}>{r['v']}</td><td>{e(r['todo'])}</td></tr>")
h.append("</table><div class=g>Газпром теплоэнерго МО 32616526616 (23,7 млн ₽, 18.11) — документация не получена.</div>")
h.append("<h2>2. Мир — лучшие тендеры, куда ISP может подать сам (маржа по спецификациям, тыс. USD)</h2><table><tr><th>Страна</th><th>Что закупают</th><th>Объём, т</th><th>Выручка</th><th>Маржа</th><th>%</th><th>Подача</th><th>Кому запросили КП</th></tr>")
asked = {"Испания": "Derypol, Bluwat", "Италия": "CarboTech, Norit (уголь); Derypol (эмульсия)", "Германия": "Steinebach, ALUCOM", "Польша": "Derypol, Bluwat", "Кипр": "FERI-TRI, Misr Chemical, CHIMI ART", "Украина": "PACHEMTECH, Lamberti", "Венгрия": "— (через Dinax/Hidrofilt)"}
for o in top:
    h.append(f"<tr><td>{e(o['country'] or '')}</td><td>{e(o['ru'][:90])}</td><td class=n>{k(o['qty_t'])}</td><td class=n>{k(o['rev_usd']/1e3)}</td><td class=n><b>{k(o['margin_usd']/1e3)}</b></td><td class=n>{o['margin_pct']:.0f}</td><td>{e(str(o['dl'])[:11])}</td><td>{e(asked.get(o['country'], '—'))}</td></tr>")
h.append("</table><div class=g>Италия: заказчик вправе отклонить товар >50% из Китая/Индии — берём европейские заводы. ЕС: антидемпинг 192% на PBTC из КНР. Африка: Уганда NWSC (сульфат, сода, ICB, 27.11) — нужен агент и банковская гарантия ≈$53,5 тыс.</div>")
h.append("<div class=cols><div><h2>3. Письма 15.11 (по группам)</h2><table><tr><th>Группа</th><th>Писем</th><th>Кому (примеры)</th></tr>")
from collections import defaultdict
gg = defaultdict(list)
for m in mail: gg[m["group"]].append(m["to"])
for g, L in gg.items(): h.append(f"<tr><td>{e(g)}</td><td class=n>{len(L)}</td><td>{e(', '.join(x.split('@')[1].split(';')[0] for x in L[:6]))}{' …' if len(L) > 6 else ''}</td></tr>")
h.append("</table></div><div><h2>4. План</h2><table><tr><th>Когда</th><th>Кто</th><th>Что</th></tr>" + "".join(f"<tr><td style='white-space:nowrap'>{a}</td><td>{e(b)}</td><td>{e(c)}</td></tr>" for a, b, c in PLAN) + "</table></div></div></body></html>")
hp = f"{SCRATCH}/itog1511.html"; open(hp, "w").write("".join(h))
subprocess.run(["/opt/pw-browsers/chromium", "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer", f"--print-to-pdf={os.path.abspath('../ITOG_1511.pdf')}", "file://" + hp], check=True, capture_output=True, timeout=180)
import pypdf; print("pages", len(pypdf.PdfReader("../ITOG_1511.pdf").pages))
H = PatternFill("solid", fgColor="0B2545"); HF = Font(bold=True, color="FFFFFF"); W = Alignment(wrap_text=True, vertical="top")
wb = Workbook()
def sheet(title, cols, rows, widths, first=False):
    ws = wb.active if first else wb.create_sheet(); ws.title = title; ws.append(cols)
    for c in ws[1]: c.fill = H; c.font = HF; c.alignment = W
    for r in rows: ws.append(r)
    for row in ws.iter_rows(min_row=2):
        for c in row: c.alignment = W
    for i, w in enumerate(widths, 1): ws.column_dimensions[CL(i)].width = w
    ws.freeze_panes = "A2"
sheet("1. РФ", ["Номер", "Заказчик", "Предмет", "Подача", "НМЦ, ₽", "Прибыль −10%", "Маржа −10%, %", "Прибыль −20%", "Решение", "Что сделать"],
      [[r["n"], r["cust"], r["what"], r["dl"], r["nmc"], r["p10"], r["m10"], r["p20"], r["v"], r["todo"]] for r in rf], [20, 24, 34, 16, 13, 13, 10, 13, 14, 70], True)
sheet("2. Мир топ", ["Страна", "Номер", "Что закупают", "Объём, т", "Выручка, USD", "Себестоимость, USD", "Маржа, USD", "Маржа, %", "Подача", "Ссылка"],
      [[o["country"], o["n"], o["ru"], o["qty_t"], o["rev_usd"], o["cost_usd"], o["margin_usd"], o["margin_pct"], o["dl"], o["link"]] for o in top], [12, 24, 60, 9, 13, 13, 13, 8, 16, 40])
sheet("3. Письма 15.11", ["Группа", "Кому", "Продукт / компания", "Тендер"], [[m["group"], m["to"], m["what"], m["tender"]] for m in mail], [28, 40, 36, 40])
sheet("4. План", ["Когда", "Кто", "Что"], PLAN, [12, 16, 110])
wb.save("../ITOG_1511.xlsx"); print("ok", len(mail))
