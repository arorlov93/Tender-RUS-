"""Итоговый Excel 15.11: все закупки со ссылками (каждая проверена), описание по-русски, потенциал и маржа по миру.
Вход: scratch/rows_1511.json (tenders_1511_table.py), linkcheck.tsv, b2b_resolved.json, ted_found.json, gp/*.json,
      data/tenders_1511/{world_*.json, ru_chem.json, mir_ru.json, fx_usd_1511.json}
→ ../ZAKUPKI_SSYLKI_I_POTENCIAL_1511.xlsx"""
import json, re, os, glob
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as CL
from isp_letterhead import SCRATCH

T = "data/tenders_1511/"
R = json.load(open(f"{SCRATCH}/rows_1511.json"))
CHK = {}
for l in open(f"{SCRATCH}/linkcheck.tsv"):
    u, rest = l.rstrip("\n").split("\t"); CHK[u] = rest.split(" ")[0]
B2B = json.load(open(f"{SCRATCH}/b2b_resolved.json"))
TED = json.load(open(f"{SCRATCH}/ted_found.json"))
FX = json.load(open(T + "fx_usd_1511.json"))["rates"]
MR = {x["id"]: x for x in json.load(open(T + "mir_ru.json"))} if os.path.exists(T + "mir_ru.json") else {}
RU = {str(t["number"]): t for t in json.load(open(T + "ru_chem.json"))["tenders"]}
AGG = {}
for fp in glob.glob(T + "world_*.json"):
    for t in json.load(open(fp)).get("tenders", []):
        k = str(t.get("tender_number") or "").strip()
        if k and t.get("aggregator_url"): AGG.setdefault(k, t["aggregator_url"])
HOME = ("bidplus.gem.gov.in/all-bids", "etpgpb.ru/", "www.rts-tender.ru/", "www.astgoz.ru/", "www.ktendering.com.kw/", "eproc.publicprocurement.govmu.org", "gsecl.nprocure.com", "tender.samolet.ru/", "steg.com.tn/fr/all-tenders")

def gp_ok(n): f = f"{SCRATCH}/gp/{n}.json"; return os.path.exists(f) and os.path.getsize(f) > 200 and "null" != open(f).read(10).strip()

def resolve(r):
    """→ (основная ссылка, статус проверки, запасная ссылка)"""
    n, u = r["n"], r.get("link")
    alt = AGG.get(n) or (RU.get(n, {}).get("source_url") if RU.get(n, {}).get("source_url") != u else None)
    if re.fullmatch(r"\d{7}", n) and n in B2B and B2B[n][0]:
        return B2B[n][0], ("Открывается ✓ (HTTP 200, 15.11)" if B2B[n][1] == "200" else f"HTTP {B2B[n][1]}"), alt
    if not u: return None, "Ссылки нет — искать по номеру на площадке", alt
    m = re.findall(r"(\d{3,7}-2026)", u) if "ted.europa.eu" in u else []
    if m:
        main = f"https://ted.europa.eu/en/notice/-/detail/{m[0]}"
        if len(m) > 1: alt = f"https://ted.europa.eu/en/notice/-/detail/{m[1]}"
        return main, ("Объявление подтверждено через API TED ✓" if all(x in TED for x in m) else "TED: номер не найден в API"), alt
    if "zakupki.gov.ru" in u:
        return u, ("Номер подтверждён по зеркалу ЕИС ✓ (zakupki.gov.ru из облака не открывается, в РФ открывается)" if gp_ok(n) else "Номер из выгрузки ЕИС; zakupki.gov.ru из облака не открывается — проверить у себя"), alt
    if any(h in u for h in HOME): return u, "Ссылка на площадку (не на саму закупку) — искать по номеру", alt
    c = CHK.get(u, "?")
    if c == "200": return u, "Открывается ✓ (HTTP 200, 15.11)", alt
    if c == "403": return u, "Сайт закрыт от роботов (403) — в браузере обычно открывается; есть запасная", alt
    if c == "404": return u, "Страница не найдена (404) — использовать запасную", alt
    if c in ("000", "500"): return u, "Из облака не открылся (ограничен зарубежный доступ) — проверить у себя; есть запасная" if alt else "Из облака не открылся — проверить у себя", alt
    return u, f"HTTP {c}", alt

H = PatternFill("solid", fgColor="0B2545"); HF = Font(bold=True, color="FFFFFF"); W = Alignment(wrap_text=True, vertical="top")
LINK = Font(color="0563C1", underline="single"); OK = Font(color="1B7F3B"); BAD = Font(color="B00020"); B = Font(bold=True)
NOTE = PatternFill("solid", fgColor="FFF4CE")
wb = Workbook()

def sheet(title, cols, widths):
    ws = wb.create_sheet(title); ws.append(cols)
    for c in ws[1]: c.fill = H; c.font = HF; c.alignment = W
    for i, w in enumerate(widths, 1): ws.column_dimensions[CL(i)].width = w
    ws.freeze_panes = "C2"; return ws
def link(ws, row, col, url):
    if url: c = ws.cell(row, col); c.value = url; c.hyperlink = url; c.font = LINK
def fin(ws):
    for row in ws.iter_rows(min_row=2):
        for c in row: c.alignment = W
    ws.auto_filter.ref = ws.dimensions

stat = {}
# --- 0. Как читать
w0 = wb.active; w0.title = "Как читать"
for line in [
    ["Итоговая таблица закупок на 15.11.2026 (подача 18.11 и позже)"],
    ["Листы:"], ["1. РФ — новые (поиск 15.11) — закупки химии, которые ещё не отправлялись. Подаёт Гидропас."],
    ["2. РФ — выгрузка 13.11 — закупки из ваших выгрузок (уже отправлялись ранее), со ссылками."],
    ["3. Мир — что это за закупки — все зарубежные тендеры для ISP с описанием на русском."],
    ["4. Мир — потенциал и маржа — предварительный расчёт: стоимость тендера в USD, наша себестоимость по ценам заводов, маржа."],
    ["5. Параметры — курсы валют, цены заводов FOB, допущения (логистика, доп. расходы, целевая маржа). Меняете здесь — пересчитывается лист 4."],
    [""], ["Колонка «Проверка ссылки»: каждая ссылка проверена 15.11 из облака:"],
    ["  ✓ HTTP 200 — страница открылась; TED — объявление подтверждено через официальный API TED; ЕИС — номер подтверждён по зеркалу ЕИС (zakupki.gov.ru из облака не открывается, в РФ открывается)."],
    ["  «Из облака не открылся» — сайт ограничивает зарубежный/облачный доступ (Турция, Тунис, Марокко, ЮАР и др.); из браузера обычно открывается. Есть запасная ссылка (агрегатор)."],
    ["  «Ссылка на площадку» — у площадки нет прямой ссылки на закупку без входа (GeM, РТС, АСТ ГОЗ, ЭТП ГПБ…): ищите по номеру."],
    [""], ["Расчёт маржи — ПРЕДВАРИТЕЛЬНЫЙ: стоимость тендера — опубликованная оценка заказчика; цены FOB — из КП заводов (Aierfuke, Kairui, Suqing); логистика и доп. расходы — ДОПУЩЕНИЯ (ставки экспедиторов по этим странам не запрашивались)."],
]: w0.append(line)
w0["A1"].font = Font(bold=True, size=13); w0.column_dimensions["A"].width = 160

# --- 1-2. РФ
def ru_sheet(title, rows):
    ws = sheet(title, ["№", "Номер", "Ссылка (основная)", "Проверка ссылки", "Запасная ссылка", "Предмет", "Заказчик / регион", "Площадка", "НМЦ, ₽", "Подача (МСК)", "Статус / что сделать", "Сделано ✓"],
               [4, 20, 50, 40, 40, 50, 36, 16, 14, 14, 40, 9])
    for i, r in enumerate(rows, 1):
        u, st, alt = resolve(r); stat[st.split(" (")[0].split(" —")[0]] = stat.get(st.split(" (")[0].split(" —")[0], 0) + 1
        ws.append([i, r["n"], None, st, None, r["subj"], r["cust"], r["plat"], r["nmc"] if isinstance(r["nmc"], (int, float)) else None, r["dl"], (r.get("st") or "") + (": " + r["todo"] if r.get("todo") else ""), ""])
        link(ws, ws.max_row, 3, u); link(ws, ws.max_row, 5, alt)
        ws.cell(ws.max_row, 4).font = OK if "✓" in st else BAD
        ws.cell(ws.max_row, 9).number_format = "# ##0"
    fin(ws)
ru_sheet("1. РФ — новые", [r for r in R if r["src"] == "РФ (поиск 15.11)"])
ru_sheet("2. РФ — выгрузка 13.11", [r for r in R if r["src"] == "РФ (выгрузка 13.11)"])

# --- 3. Мир — описание
mir = [r for r in R if r["who"] == "ISP"]
for i, r in enumerate(mir): r["_id"] = i; r["_m"] = MR.get(i, {})
order = {"Напрямую (иностранцы допускаются)": 0, "Уточнить": 1, "Через местного партнёра / регистрацию": 2}
mir.sort(key=lambda r: (order.get(r["_m"].get("access_ru"), 1), r["d"]))
ws = sheet("3. Мир — что это за закупки", ["№", "Доступ для ISP", "Страна", "Заказчик", "Номер", "Что закупают (по-русски)", "Группа товара", "Объём, т", "Подача", "Ссылка (основная)", "Проверка ссылки", "Запасная ссылка", "Условия допуска / документы", "Сделано ✓"],
           [4, 22, 14, 34, 20, 60, 30, 10, 16, 48, 36, 40, 50, 9])
for i, r in enumerate(mir, 1):
    m = r["_m"]; u, st, alt = resolve(r); k = st.split(" (")[0].split(" —")[0]; stat[k] = stat.get(k, 0) + 1
    ws.append([i, m.get("access_ru") or r["acc"], m.get("country_ru"), r["cust"], r["n"], m.get("ru") or r["subj"], m.get("group"), m.get("qty_t"), r["dl"], None, st, None, (r["note"] or "") + (" | " + r["todo"] if r["todo"] else ""), ""])
    link(ws, ws.max_row, 10, u); link(ws, ws.max_row, 12, alt)
    ws.cell(ws.max_row, 11).font = OK if "✓" in st else BAD
fin(ws)

# --- 5. Параметры
wp = sheet("5. Параметры", ["Параметр", "Значение", "Ед.", "Источник / комментарий"], [52, 14, 12, 90])
P = {}
def par(name, val, unit, src):
    wp.append([name, val, unit, src]); P[name] = f"'5. Параметры'!$B${wp.max_row}"
    if "ДОПУЩЕНИЕ" in src: wp.cell(wp.max_row, 2).fill = NOTE
par("Логистика до объекта (море + доставка), USD/т", 150, "USD/т", "ДОПУЩЕНИЕ: ставки экспедиторов по этим странам не запрашивались. Для ориентира: авто Китай→Каракалпакстан $17 760 / 22 т ≈ $807/т (Saada Log, 13.11)")
par("Доп. расходы (пошлины, банк, страховка, документы), % к FOB+логистике", 0.08, "доля", "ДОПУЩЕНИЕ")
par("Целевая маржа, если нет объёма (доля от стоимости тендера)", 0.12, "доля", "ДОПУЩЕНИЕ — используется только когда объём не опубликован")
wp.append([]); wp.append(["Цены заводов FOB (из КП)", None, None, None]); wp.cell(wp.max_row, 1).font = B
PRICE = {"PAC (полиоксихлорид алюминия)": (295, "Henan Aierfuke, PAC порошок технический FOB Qingdao (питьевой $355). Если в ТЗ жидкий PAC — цену пересчитать по Al2O3"),
         "Фосфонаты": (1315, "Kairui HEDP 60% FOB Qingdao 11.11 (ATMP 50% $1 195; PBTCA 50% $1 345)"),
         "Антискалант/ингибитор накипи и коррозии (водоподготовка)": (1345, "ориентир: Kairui PBTCA 50% FOB $1 345 — готовый антискалант-рецептуру Kairui не котировал"),
         "Ионообменные смолы": (None, "Suqing FOB Shanghai: катионит 001x7 $840/м³, анионит 201x7 $2 090/м³ — цена за м³, не за т; считать вручную"),
         "Сульфат алюминия": (None, "КП завода нет"), "Хлорное/сернокислое железо": (None, "КП завода нет"), "Флокулянт (полиакриламид/полиэлектролит)": (None, "КП завода нет"),
         "Деэмульгатор": (None, "Rimpro: цена после лаб. подбора"), "Ингибитор коррозии (нефтегаз)": (None, "Rimpro: цена после лаб. подбора"),
         "Активированный уголь": (None, "КП завода нет"), "Оксид алюминия / молекулярные сита / адсорбенты": (None, "Jalon: КП ждём")}
for g, (v, s) in PRICE.items(): par("FOB, USD/т — " + g, v, "USD/т", s)
wp.append([]); wp.append(["Курсы валют (за 1 USD)", None, None, "open.er-api.com (ExchangeRate-API), обновлено " + json.load(open(T + "fx_usd_1511.json"))["time_last_update_utc"]]); wp.cell(wp.max_row, 1).font = B
FXROW = {}
for cur in sorted({(x.get("value_currency") or "") for x in MR.values()} - {""}):
    if cur in FX: wp.append([cur, FX[cur], "за 1 USD", None]); FXROW[cur] = f"'5. Параметры'!$B${wp.max_row}"
fin(wp)

# --- 4. Потенциал и маржа
wm = sheet("4. Мир — потенциал и маржа", ["№", "Доступ", "Страна", "Номер", "Что закупают", "Группа", "Объём, т", "Стоимость тендера", "Валюта", "Стоимость, USD", "Наша цена FOB, USD/т", "Себестоимость с логистикой, USD", "Маржа, USD", "Маржа, %", "Как посчитано", "Подача", "Ссылка"],
           [4, 18, 14, 18, 50, 26, 10, 16, 8, 14, 12, 16, 14, 9, 50, 14, 40])
wm.insert_rows(2); wm.cell(2, 1).value = "ИТОГО"; wm.cell(2, 1).font = B
rows4 = [r for r in mir if (r["_m"].get("value_amount") or r["_m"].get("qty_t")) and r["_m"].get("group") != "Прочее / не наш профиль"]
def acc_rank(r): return order.get(r["_m"].get("access_ru"), 1)
rows4.sort(key=lambda r: (acc_rank(r), -(r["_m"].get("value_amount") or 0) / (FX.get(r["_m"].get("value_currency") or "USD") or 1)))
for i, r in enumerate(rows4, 1):
    m = r["_m"]; u, st, alt = resolve(r); x = wm.max_row + 1
    g = m.get("group"); pk = "FOB, USD/т — " + str(g)
    wm.append([i, m.get("access_ru"), m.get("country_ru"), r["n"], m.get("ru") or r["subj"], g, m.get("qty_t"), m.get("value_amount"), m.get("value_currency")])
    cur = m.get("value_currency")
    if m.get("value_amount") and cur:
        wm.cell(x, 10).value = f"=H{x}" if cur == "USD" else (f"=H{x}/{FXROW[cur]}" if cur in FXROW else None)
    if pk in P and PRICE.get(g, (None,))[0]: wm.cell(x, 11).value = f"={P[pk]}"
    else: wm.cell(x, 11).value = "нет КП завода"
    if m.get("qty_t") and pk in P and PRICE.get(g, (None,))[0]:
        wm.cell(x, 12).value = f"=G{x}*(K{x}+{P['Логистика до объекта (море + доставка), USD/т']})*(1+{P['Доп. расходы (пошлины, банк, страховка, документы), % к FOB+логистике']})"
        if wm.cell(x, 10).value:
            wm.cell(x, 13).value = f"=J{x}-L{x}"; wm.cell(x, 14).value = f"=IF(J{x}>0,M{x}/J{x},\"\")"; how = "стоимость тендера − (объём × (FOB + логистика)) × (1 + доп. расходы)"
        else:
            how = "стоимость не опубликована: в колонке L — наша себестоимость = минимальная цена предложения; цена с целевой маржой = L / (1 − целевая маржа)"
            wm.cell(x, 13).value = f"=L{x}/(1-{P['Целевая маржа, если нет объёма (доля от стоимости тендера)']})-L{x}"
            wm.cell(x, 14).value = f"={P['Целевая маржа, если нет объёма (доля от стоимости тендера)']}"
    elif wm.cell(x, 10).value:
        wm.cell(x, 13).value = f"=J{x}*{P['Целевая маржа, если нет объёма (доля от стоимости тендера)']}"; wm.cell(x, 14).value = f"={P['Целевая маржа, если нет объёма (доля от стоимости тендера)']}"
        how = "нет объёма или цены завода: маржа = стоимость × целевая маржа (оценка потенциала)"
    else:
        how = "есть объём, но нет цены завода и стоимости — нужна КП завода"
    risk = {"Антискалант/ингибитор накипи и коррозии (водоподготовка)": " ⚠ себестоимость по сырью PBTCA, готовая рецептура дороже — маржа завышена",
            "PAC (полиоксихлорид алюминия)": " ⚠ у нас PAC-порошок; если в ТЗ жидкий PAC (~10% Al2O3) — тоннаж и цену пересчитать",
            "Пакет разной химии": " ⚠ пакет: часть позиций не наш профиль"}.get(g, "")
    if m.get("value_amount") and not cur: risk += " ⚠ валюта не указана в источнике"
    wm.cell(x, 15).value = how + risk; wm.cell(x, 16).value = r["dl"]; link(wm, x, 17, u)
    for c in (8, 10, 12, 13): wm.cell(x, c).number_format = "# ##0"
    wm.cell(x, 14).number_format = "0%"
last = wm.max_row
for c in (10, 12, 13): wm.cell(2, c).value = f"=SUM({CL(c)}3:{CL(c)}{last})"; wm.cell(2, c).number_format = "# ##0"; wm.cell(2, c).font = B
for j, a in enumerate(["Напрямую (иностранцы допускаются)", "Уточнить", "Через местного партнёра / регистрацию"]):
    rr = last + 2 + j
    wm.cell(rr, 2).value = a; wm.cell(rr, 5).value = "Подытог по группе доступа"; wm.cell(rr, 2).font = B
    for c in (10, 12, 13): wm.cell(rr, c).value = f"=SUMIF($B$3:$B${last},$B{rr},{CL(c)}3:{CL(c)}{last})"; wm.cell(rr, c).number_format = "# ##0"; wm.cell(rr, c).font = B
wm.cell(2, 5).value = f"Строк: {len(rows4)} (из {len(mir)} — только где есть стоимость или объём и товар нашего профиля)"
fin(wm)
wb.move_sheet("5. Параметры", offset=0)
wb.save("../ZAKUPKI_SSYLKI_I_POTENCIAL_1511.xlsx")
print("rows ru_new", sum(1 for r in R if r["src"] == "РФ (поиск 15.11)"), "ru_old", sum(1 for r in R if r["src"] == "РФ (выгрузка 13.11)"), "mir", len(mir), "margin rows", len(rows4))
print(stat)

# ---------- разделение: РФ (Excel) и Мир (Excel + PDF)
from openpyxl import load_workbook
src = "../ZAKUPKI_SSYLKI_I_POTENCIAL_1511.xlsx"
a = load_workbook(src)
for n in ["5. Параметры", "3. Мир — что это за закупки", "4. Мир — потенциал и маржа"]: del a[n]
a["Как читать"]["A1"].value = "Российские закупки (Гидропас) на 15.11.2026, подача 18.11 и позже — ссылки проверены"
for row in list(a["Как читать"].iter_rows(min_row=2)):
    if row[0].value and any(k in str(row[0].value) for k in ("Мир", "Параметры", "маржи")): row[0].value = None
a.save("../ZAKUPKI_RF_SSYLKI_1511.xlsx")
b = load_workbook(src)
for n in ["1. РФ — новые", "2. РФ — выгрузка 13.11"]: del b[n]
b["Как читать"]["A1"].value = "Зарубежные закупки (ISP Group) на 15.11.2026, подача 18.11 и позже — описание, ссылки, потенциал и маржа"
for row in list(b["Как читать"].iter_rows(min_row=2)):
    if row[0].value and "РФ" in str(row[0].value): row[0].value = None
b.save("../ZAKUPKI_MIR_1511.xlsx")
os.remove(src)

# ---------- PDF по миру: описание + примерная маржа (расчёт в Python по тем же параметрам)
import html, subprocess
LOG, EXTRA, TGT = 150, 0.08, 0.12
calc = []
for r in rows4:
    m = r["_m"]; g = m.get("group"); cur = m.get("value_currency"); amt = m.get("value_amount"); q = m.get("qty_t")
    usd = (amt if cur == "USD" else amt / FX[cur]) if amt and cur and (cur == "USD" or cur in FX) else None
    fob = PRICE.get(g, (None,))[0]
    cost = q * (fob + LOG) * (1 + EXTRA) if q and fob else None
    if usd and cost: mar, how = usd - cost, "расчёт"
    elif cost: mar, how = cost / (1 - TGT) - cost, "наша мин. цена"
    elif usd: mar, how = usd * TGT, "≈12% от стоимости"
    else: mar, how = None, "нет данных"
    calc.append((r, usd, cost, mar, how))
e = html.escape; fm = lambda v: "—" if v is None else f"{v/1e3:,.0f}".replace(",", " ")
grp = {}
for r, usd, cost, mar, how in calc:
    k = r["_m"].get("access_ru") or "Уточнить"; s_ = grp.setdefault(k, [0, 0, 0]); s_[0] += 1; s_[1] += usd or 0; s_[2] += mar or 0
css = ("@page{size:A4 landscape;margin:7mm}body{font-family:'DejaVu Sans',Arial;font-size:6.6px;color:#111}h1{font-size:12.5px;margin:0 0 2px;color:#0B2545}h2{font-size:9px;margin:4px 0 1px;color:#0B2545}"
       "table{border-collapse:collapse;width:100%}th{background:#0B2545;color:#fff;padding:1px 3px;text-align:left}td{border-bottom:1px solid #ddd;padding:1px 3px;vertical-align:top;line-height:1.15}"
       ".n{text-align:right;white-space:nowrap}.g{color:#555}a{color:#0563C1;text-decoration:none}.w{color:#9a6700}")
h = [f"<!DOCTYPE html><html><head><meta charset=utf-8><style>{css}</style></head><body><h1>Зарубежные тендеры ISP Group (подача от 18.11.2026): что закупают и примерная маржа</h1>",
     f"<div class=g>Из {len(mir)} найденных тендеров здесь {len(rows4)} — по которым опубликована стоимость или объём и товар нашего профиля. Полный список и ссылки — ZAKUPKI_MIR_1511.xlsx. Суммы в тыс. USD (курс ExchangeRate-API на 15.11).</div>",
     "<table style='width:60%;margin:3px 0'><tr><th>Доступ для ISP</th><th>Тендеров</th><th>Стоимость, тыс. USD</th><th>Примерная маржа, тыс. USD</th></tr>"]
for k in ["Напрямую (иностранцы допускаются)", "Уточнить", "Через местного партнёра / регистрацию"]:
    if k in grp: h.append(f"<tr><td>{e(k)}</td><td class=n>{grp[k][0]}</td><td class=n>{fm(grp[k][1])}</td><td class=n><b>{fm(grp[k][2])}</b></td></tr>")
tot = [sum(v[i] for v in grp.values()) for i in range(3)]
h.append(f"<tr><td><b>Итого</b></td><td class=n><b>{tot[0]}</b></td><td class=n><b>{fm(tot[1])}</b></td><td class=n><b>{fm(tot[2])}</b></td></tr></table>")
h.append("<table><tr><th>#</th><th>Доступ</th><th>Страна</th><th>Номер (ссылка)</th><th>Что закупают</th><th>Объём, т</th><th>Стоимость, тыс. USD</th><th>Себест., тыс. USD</th><th>Маржа, тыс. USD</th><th>Как</th><th>Подача</th></tr>")
for i, (r, usd, cost, mar, how) in enumerate(calc, 1):
    m = r["_m"]; u, st, alt = resolve(r); u = u or alt
    num = f"<a href='{e(u)}'>{e(r['n'][:22])}</a>" if u else e(r["n"][:22])
    acc = {"Напрямую (иностранцы допускаются)": "напрямую", "Через местного партнёра / регистрацию": "партнёр"}.get(m.get("access_ru"), "уточнить")
    warn = " ⚠" if m.get("group") in ("Антискалант/ингибитор накипи и коррозии (водоподготовка)", "PAC (полиоксихлорид алюминия)", "Пакет разной химии") else ""
    h.append(f"<tr><td>{i}</td><td>{acc}</td><td>{e(m.get('country_ru') or '')}</td><td>{num}</td><td>{e((m.get('ru') or r['subj'])[:150])}</td><td class=n>{'' if not m.get('qty_t') else f'{m['qty_t']:,.0f}'.replace(',', ' ')}</td>"
             f"<td class=n>{fm(usd)}</td><td class=n>{fm(cost)}</td><td class=n><b>{fm(mar)}</b></td><td class=w>{how}{warn}</td><td style='white-space:nowrap'>{e(str(r['dl'])[:11])}</td></tr>")
h.append("</table><div class=g style='margin-top:3px'>Как считано: «расчёт» — стоимость тендера минус себестоимость (объём × (FOB завода + логистика $150/т) × 1,08). «Наша мин. цена» — стоимость не опубликована, маржа при цене на 12% выше себестоимости. "
         "«≈12%» — нет цены завода (сульфат алюминия, железо, флокулянты, деэмульгаторы, уголь): оценка потенциала. Цены FOB: PAC $295/т (Aierfuke), PBTCA $1 345/т (Kairui). Логистика и 8% доп. расходов — ДОПУЩЕНИЯ. "
         "⚠ — PAC: у нас порошок, при жидком PAC пересчитать; антискалант: себестоимость по сырью, маржа завышена; пакет: часть позиций не наш профиль. Без Ирана, Сирии, OFAC SDN.</div></body></html>")
hp = f"{SCRATCH}/mir_marzha.html"; open(hp, "w").write("".join(h))
subprocess.run(["/opt/pw-browsers/chromium", "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer", f"--print-to-pdf={os.path.abspath('../ZAKUPKI_MIR_MARZHA_1511.pdf')}", "file://" + hp], check=True, capture_output=True, timeout=180)
import pypdf; print("PDF pages", len(pypdf.PdfReader("../ZAKUPKI_MIR_MARZHA_1511.pdf").pages), "totals", [round(x) for x in tot])
