"""Тендеры с подачей ≥18.11.2026 (на 15.11): ссылки для запроса ТЗ + что за нами.
Источники: ZAPROSY_KP_1311 (лист 1, 65 лотов 13.11), международные (T-47, Туркменгаз), data/tenders_1511/ru_chem.json и world_chem.json (если есть).
→ ../TENDERY_TZ_1511.xlsx + .pdf"""
import json, re, os, html, subprocess
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from isp_letterhead import SCRATCH

T = "data/tenders_1511/"
MIN = (11, 18)  # (месяц, день) — подача не раньше 18.11
def dk(s):
    s = str(s or "")
    m = re.search(r"(\d{4})-(\d\d)-(\d\d)", s)
    if m: return (int(m[2]), int(m[3])) if m[1] == "2026" else ((13, 0) if int(m[1]) > 2026 else (0, 0))
    m = re.search(r"(\d\d)\.(\d\d)(?:\.(\d{4}))?", s)
    if m:
        if m[3] and int(m[3]) > 2026: return (13, 0)
        return (int(m[2]), int(m[1]))
    return None
def eis(n):
    if re.fullmatch(r"\d{11}", n): return f"https://zakupki.gov.ru/epz/order/notice/notice223/common-info.html?regNumber={n}"
    if re.fullmatch(r"\d{19}", n): return f"https://zakupki.gov.ru/epz/order/extendedsearch/results.html?searchString={n}"
PLAT = {"АСТ ГОЗ": "https://www.astgoz.ru/", "ЭТП ГПБ": "https://etpgpb.ru/", "РТС-тендер": "https://www.rts-tender.ru/"}

rows = []
# 1) лоты 13.11
src = load_workbook("../ZAPROSY_KP_1311.xlsx")["1. Закупки и ссылки"]
for c in src.iter_rows(min_row=2):
    v = [x.value for x in c]
    n = str(v[2]).split("/")[-1]; plat = str(v[5]); term = v[8]
    if isinstance(term, (int, float)) and term > 6: continue
    d = dk(v[7])
    if not d or d < MIN: continue
    b2b = f"https://www.b2b-center.ru/market/view.html?id={n}" if re.fullmatch(r"\d{7}", n) and "b2b" in plat.lower() else None
    link = b2b or (PLAT["АСТ ГОЗ"] if "АСТ ГОЗ" in plat else None) or eis(n) or next((u for k, u in PLAT.items() if k in plat), None) or (c[9].hyperlink.target if c[9].hyperlink else None)
    st = str(v[1])
    todo = ("Запросить документацию на АСТ ГОЗ (3 раб. дня)" if "АСТ ГОЗ" in plat else "Скачать ТЗ" if st == "Запросить ТЗ" else
            "ТЗ есть — нужны КП заводов (ZAPROSY_KP_1311)" if "КП" in st or v[10] else "ТЗ есть — готовить заявку")
    rows.append(dict(src="РФ (выгрузка 13.11)", who="Гидропас", n=n, subj=str(v[3]), cust=str(v[4]), plat=plat, nmc=v[6], cur="₽", dl=str(v[7]), d=d, link=link, st=st, todo=todo, note=str(v[11] or "")[:200]))
# 2) международные, уже известные
for x in [
    dict(n="T-47/2026", subj="Химреагенты водоподготовки (11 позиций, полный пакет): антискаланты RO и градирни, бисульфит, биоциды, очистители мембран, амин, поглотитель O2, ТНФ + пассивация",
         cust="СП Uz-Kor Gas Chemical (Устюртский ГХК), Узбекистан", plat="конверт, Ташкент", dl="02.12.2026 15:00 (Ташкент)", link=None, st="Участвуем",
         todo="ТД есть. Нужны: КП Kairui/Taihe, свой специалист на апробацию 1 мес, банковская гарантия 8 000 USD, доверенность завода", note="Срок подтвердил А. Касимов 13.11"),
    dict(n="T/GAZ-038 Лот 5", subj="Химическая продукция: SY-210 (14 т), ингибитор коррозии «Селерин» 150,56 т, антивспениватели, метанол, NH4HF2, буровая химия",
         cust="ГК «Туркменгаз», Туркменистан", plat="turkmengaz.gov.tm", dl="23.11.2026 15:00", link="https://turkmengaz.gov.tm/competition-detail/209", st="Ранее решили пропустить",
         todo="Пересмотреть: есть LOA Kairui и Rimpro (эксклюзив ТМ). Взнос 1 500 USD, образцы, регистрация в госреестре ТМ", note="ТД: https://turkmengaz.gov.tm/competition-download/209"),
]:
    x.update(src="Международные (ранее)", who="ISP", nmc=None, cur="", d=dk(x["dl"])); rows.append(x)
# 3) находки агентов 15.11
def load(f):
    try: return json.load(open(T + f)).get("tenders", [])
    except FileNotFoundError: return []
for t in load("ru_chem.json"):
    d = dk(t.get("deadline") or t.get("application_deadline") or t.get("deadline_msk"))
    if not d or d < MIN: continue
    rows.append(dict(src="РФ (поиск 15.11)", who="Гидропас", n=str(t.get("number") or t.get("num")), subj=str(t.get("subject") or t.get("title") or ""), cust=f"{t.get('customer') or ''} / {t.get('region') or ''}",
                     plat=str(t.get("platform") or ""), nmc=t.get("nmc") or t.get("nmc_rub"), cur="₽", dl=str(t.get("deadline") or t.get("application_deadline") or t.get("deadline_msk")), d=d,
                     link=(t.get("source_url") if "b2b-center" in str(t.get("link") or "") else None) or t.get("link") or t.get("url") or t.get("source_url"), st="Новая", todo="Скачать ТЗ" + (" (ПП 1875: " + str(t.get("pp1875")) + ")" if t.get("pp1875") else ""),
                     note=str(t.get("matches") or t.get("product_group") or "")[:200]))
import glob
WT = []
for fp in sorted(glob.glob(T + "world_*.json")):
    for t in load(os.path.basename(fp)): t["_reg"] = os.path.basename(fp)[6:-5]; WT.append(t)
seen = set()
for t in WT:
    key = (str(t.get("tender_number") or "").strip().lower(), str(t.get("buyer") or "")[:20].lower())
    if key[0] and key in seen: continue
    if re.search(r"T/GAZ-03[78]|T-47", key[0], re.I): continue  # уже есть в «Международные (ранее)» / не профиль
    seen.add(key)
    d = dk(t.get("deadline"))
    if not d or d < MIN: continue
    rows.append(dict(src="Мир 15.11: " + t["_reg"], who="ISP", n=str(t.get("tender_number") or t.get("number") or ""), subj=str(t.get("title") or "") + (" — " + str(t.get("products_quantities")) if t.get("products_quantities") else ""), cust=f"{t.get('buyer') or ''} / {t.get('country') or ''}",
                     plat=str(t.get("portal") or ""), nmc=t.get("estimated_value"), cur="", dl=str(t.get("deadline")), d=d, link=t.get("url") or t.get("notice_url") or t.get("aggregator_url"), st="Новая",
                     todo=str(t.get("how_to_get_documents") or t.get("documents") or "Получить ТД")[:160], note=str(t.get("eligibility") or "")[:200]))
for r in rows:
    v = r["nmc"]
    if v is not None and not isinstance(v, (int, float, str)): r["nmc"] = json.dumps(v, ensure_ascii=False)[:120]
def access(r):
    if r["who"] == "Гидропас": return "Гидропас (РФ)"
    t = (r.get("note", "") + " " + r.get("todo", "")).lower()
    if re.search(r"only local|local only|yerli|sadece|national competitive|\bncb\b|local registration|registered in tunisia|mauritian|cnpj|indian seller|make in india|class-i|\bgem\b|local supplier|bahrain|etimad|saudi cr|local presence|только мест|местн", t): return "Через партнёра / местная регистрация"
    if re.search(r"international|\bicb\b|foreign|open to all|eu-wide|\bted\b|open procedure|иностран|международ", t): return "Напрямую (возможно)"
    return "Уточнить"
for r in rows: r["acc"] = access(r)
rows.sort(key=lambda r: (r["d"], r["src"]))

wb = Workbook(); ws = wb.active; ws.title = "Тендеры ≥18.11"
H = PatternFill("solid", fgColor="0B2545"); HF = Font(bold=True, color="FFFFFF"); W = Alignment(wrap_text=True, vertical="top"); LINK = Font(color="0563C1", underline="single")
cols = ["№", "Доступ для нас", "Кто подаёт", "Источник", "Номер", "Предмет", "Заказчик", "Площадка", "НМЦ", "Подача", "Ссылка", "Статус", "Что сделать", "Примечание", "Готово (✓)"]
ws.append(cols)
for c in ws[1]: c.fill = H; c.font = HF; c.alignment = W
for i, r in enumerate(rows, 1):
    ws.append([i, r["acc"], r["who"], r["src"], r["n"], r["subj"], r["cust"], r["plat"], r["nmc"], r["dl"], r["link"] or "искать по номеру на площадке", r["st"], r["todo"], r["note"], ""])
    if r["link"]: c = ws.cell(ws.max_row, 11); c.hyperlink = r["link"]; c.font = LINK
    if isinstance(r["nmc"], (int, float)): ws.cell(ws.max_row, 9).number_format = "# ##0"
for row in ws.iter_rows(min_row=2):
    for c in row: c.alignment = W
for col, w in zip("ABCDEFGHIJKLMNO", [4, 18, 10, 14, 20, 44, 34, 16, 13, 14, 42, 14, 40, 40, 9]): ws.column_dimensions[col].width = w
ws.freeze_panes = "E2"; ws.auto_filter.ref = ws.dimensions
P = json.load(open(T + "pending_1511.json"))
w2 = wb.create_sheet("Что за нами"); w2.append(["Приоритет", "Срок", "Кто", "Что сделать", "Зачем", "Статус"])
for x in P["on_us"]: w2.append([x["prio"], x["deadline"], x["who"], x["what"], x["why"], x["status"]])
w3 = wb.create_sheet("Ждём ответа"); w3.append(["От кого", "Что"])
for x in P["waiting"]: w3.append([x["who"], x["what"]])
for s, ws_ in ((w2, [9, 10, 16, 70, 40, 30]), (w3, [30, 90])):
    for c in s[1]: c.fill = H; c.font = HF
    for row in s.iter_rows(min_row=2):
        for c in row: c.alignment = W
    for i, w in enumerate(ws_): s.column_dimensions[chr(65 + i)].width = w
wb.save("../TENDERY_TZ_1511.xlsx")

e = html.escape; f = lambda v: f"{v:,.0f}".replace(",", " ") if isinstance(v, (int, float)) else "—"
css = ("@page{size:A4 landscape;margin:7mm}body{font-family:'DejaVu Sans',Arial;font-size:7.2px;color:#111}h1{font-size:13px;margin:0 0 2px;color:#0B2545}h2{font-size:9.5px;margin:5px 0 1px;color:#0B2545}"
       "table{border-collapse:collapse;width:100%}th{background:#0B2545;color:#fff;padding:2px 3px;text-align:left}td{border-bottom:1px solid #ddd;padding:1px 3px;vertical-align:top}.n{text-align:right;white-space:nowrap}.g{color:#555}a{color:#0563C1;text-decoration:none}")
PR = [r for r in rows if r["acc"] == "Напрямую (возможно)" or r["src"] == "Международные (ранее)"]
h = [f"<!DOCTYPE html><html><head><meta charset=utf-8><style>{css}</style></head><body><h1>Тендеры с подачей 18.11 и позже: ISP — куда можно подать напрямую ({len(PR)} из {len(rows)})</h1>",
     "<div class=g>Полный список ({len(rows)}: РФ/Гидропас, через партнёра, уточнить) — в Excel. Номер — ссылка. Порядок — по дате подачи.</div><table><tr><th>#</th><th>Кто</th><th>Номер</th><th>Предмет</th><th>Заказчик</th><th>НМЦ</th><th>Подача</th><th>Статус</th><th>Что сделать</th></tr>"]
for i, r in enumerate(PR, 1):
    num = f"<a href='{e(r['link'])}'>{e(r['n'])}</a>" if r["link"] else e(r["n"])
    h.append(f"<tr><td>{i}</td><td>{e(r['who'])}</td><td>{num}</td><td>{e(r['subj'][:70])}</td><td>{e(r['cust'][:42])}</td><td class=n>{f(r['nmc'])}{(' ' + r['cur']) if isinstance(r['nmc'], (int, float)) else ''}</td>"
             f"<td style='white-space:nowrap'>{e(r['dl'][:17])}</td><td>{e(r['st'])}</td><td>{e(r['todo'][:80])}</td></tr>")
h.append("</table><div class=g>Листы Excel: «Что за нами», «Ждём ответа».</div></body></html>")
hp = f"{SCRATCH}/t1511.html"; open(hp, "w").write("".join(h))
subprocess.run(["/opt/pw-browsers/chromium", "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer", f"--print-to-pdf={os.path.abspath('../TENDERY_TZ_1511.pdf')}", "file://" + hp], check=True, capture_output=True, timeout=180)
from collections import Counter
print(len(rows), Counter(r["src"] for r in rows), Counter(r["st"] for r in rows))
json.dump([{k: v for k, v in r.items() if k != "d"} | {"d": list(r["d"])} for r in rows], open(f"{SCRATCH}/rows_1511.json", "w"), ensure_ascii=False)
