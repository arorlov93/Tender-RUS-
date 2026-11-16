"""США: сводка госзакупок водоочистной химии (16.11.2026) → ../USA_ZAKUPKI_1611.xlsx + .pdf
Источники: data/usa/{municipal_south,municipal_rest,federal,market_model}_1611.json"""
import json, re, os, html, subprocess
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as CL
from isp_letterhead import SCRATCH
U = "data/usa/"
S = json.load(open(U + "municipal_south_1611.json")); R = json.load(open(U + "municipal_rest_1611.json"))
F = json.load(open(U + "federal_1611.json")); M = json.load(open(U + "market_model_1611.json"))
txt = lambda v: v if isinstance(v, str) else "; ".join((f"{x.get('item')} {x.get('qty') or ''} {x.get('unit') or ''}".strip() if isinstance(x, dict) else str(x)) for x in v) if isinstance(v, list) else json.dumps(v, ensure_ascii=False) if v else ""
# категории: (regex, метка, статус для ISP)
CAT = [(r"activated carbon|\bgac\b|powdered activated|activated.{0,10}\(pac\)|lignite", "уголь (GAC/PAC)", "GO"),
       (r"permanganate|kmno4|namno4", "перманганат", "GO"),
       (r"orthophosph|polyphosph|phosphate|corrosion inhib|scale inhib|zinc", "фосфаты/ингибиторы", "GO?"),
       (r"polymer|polyacrylamide|emulsion", "полимеры", "GO?"),
       (r"fluosilic|hydrofluosilic|\bhfs\b|fluoride", "HFS/фторид", "GO?"),
       (r"salt|nacl", "соль", "?"),
       (r"pacl|polyalumin|aluminum chlorohydrate|\bach\b|delpac", "PACl/ACH", "NO"),
       (r"hypochlorite|bleach|naocl", "гипохлорит", "NO"), (r"caustic|sodium hydroxide|naoh", "каустик", "NO"),
       (r"chlorine", "хлор", "NO"), (r"lime|quicklime|hydrated", "известь", "NO"), (r"\bco2\b|carbon dioxide|dry ice", "CO2/лёд", "NO"),
       (r"alum|aluminum sulfate", "сульфат Al", "NO"), (r"ferric|ferrous", "соли железа", "NO"), (r"bisulfite", "бисульфит", "NO")]
def cats(t):
    t = t.lower(); out = []
    for rx, lab, st in CAT:
        if re.search(rx, t) and lab not in [o[0] for o in out]:
            if lab == "уголь (GAC/PAC)" and re.search(r"carbon dioxide", t) and not re.search(r"activated|gac|powdered", t): continue
            out.append((lab, st))
    return out
def verdict(cs):
    go = [l for l, s in cs if s == "GO"]; mb = [l for l, s in cs if s in ("GO?", "?")]
    if go: return "ДА — подаём по позициям: " + ", ".join(go + mb), 0
    if mb: return "ВОЗМОЖНО — позиции: " + ", ".join(mb) + " (нужны КП)", 1
    return "НЕТ — только товарная химия местных заводов", 2
def due_key(s):
    m = re.search(r"(20\d\d)-(\d\d)-(\d\d)", s or "") or re.search(r"(\d\d)\.(\d\d)", s or "")
    if not m: return "9999"
    return f"{m[1]}-{m[2]}-{m[3]}" if len(m.groups()) == 3 else f"2026-{m[2]}-{m[1]}"
ROWS = []
for b in S["bids"]:
    t = txt(b.get("products")) + " " + (b.get("title") or "")
    ROWS.append(dict(src="Юг", st=b.get("state"), ent=b.get("entity"), num=b.get("bid_number"), title=b.get("title"), prod=txt(b.get("products"))[:600],
                     qty=txt(b.get("quantities"))[:300], val=b.get("estimated_value_usd"), due=f"{b.get('due')} {b.get('tz') or ''}", sub=txt(b.get("submission"))[:120],
                     req="; ".join(x for x in [txt(b.get("bonds")), txt(b.get("insurance")), txt(b.get("standards")), txt(b.get("local_preference")), txt(b.get("prebid"))] if x)[:500],
                     link=b.get("link"), fit=b.get("fit_score"), cats=cats(t)))
for b in R["bids"]:
    if str(b.get("status", "")).lower().startswith(("closed", "past")): continue
    t = txt(b.get("products")) + " " + (b.get("title") or "") + " " + str(b.get("category") or "")
    ROWS.append(dict(src="Остальные штаты", st=b.get("state"), ent=b.get("entity"), num=b.get("bid_number"), title=b.get("title"), prod=txt(b.get("products"))[:600],
                     qty="", val=b.get("estimated_value_usd"), due=f"{b.get('due_date')} {b.get('due_time') or ''} {b.get('tz') or ''}", sub=txt(b.get("submission"))[:120],
                     req="; ".join(x for x in [txt(b.get("bonds")), txt(b.get("insurance")), txt(b.get("nsf_awwa")), txt(b.get("local_preference")), txt(b.get("notes"))] if x)[:500],
                     link=b.get("link"), fit=b.get("fit_score_1_5"), cats=cats(t)))
for b in F["opportunities"]:
    t = txt(b.get("products")) + " " + (b.get("title") or "")
    ROWS.append(dict(src="Федеральные", st=b.get("place_of_delivery"), ent=b.get("agency"), num=b.get("solicitation_number"), title=b.get("title"), prod=txt(b.get("products"))[:600],
                     qty=txt(b.get("quantity_or_value"))[:300], val=None, due=b.get("response_deadline"), sub="SAM.gov",
                     req="; ".join(x for x in [txt(b.get("set_aside")), txt(b.get("buy_american_taa")), txt(b.get("isp_eligibility")), txt(b.get("nsf_awwa"))] if x)[:500],
                     link=b.get("link"), fit=b.get("feasibility_score_1to5"), cats=cats(t)))
for r in ROWS:
    r["v"], r["vk"] = verdict(r["cats"])
    if r["src"] == "Федеральные" and re.search(r"SDVOSB|service-disabled|8\(a\)|HUBZone|WOSB", r["req"], re.I): r["v"], r["vk"] = "НЕТ — квота недоступна ISP", 3
    r["dk"] = due_key(str(r["due"]))
ROWS = [r for r in ROWS if r["dk"] >= "2026-11-20"]
ROWS.sort(key=lambda r: (r["vk"], r["dk"]))
H = PatternFill("solid", fgColor="0B2545"); HF = Font(bold=True, color="FFFFFF"); W = Alignment(wrap_text=True, vertical="top"); LINK = Font(color="0563C1", underline="single")
FILL = {0: PatternFill("solid", fgColor="E6F4EA"), 1: PatternFill("solid", fgColor="FFF4D6"), 2: PatternFill("solid", fgColor="F3F3F3"), 3: PatternFill("solid", fgColor="FCE8E8")}
wb = Workbook(); ws = wb.active; ws.title = "Закупки США"
cols = ["Решение для ISP", "Источник", "Штат/место", "Заказчик", "Номер", "Предмет", "Продукты", "Объём", "Оценка $/год", "Подача", "Как подавать", "Требования (залоги, страховки, NSF, преференции)", "Оценка источника 1–5", "Ссылка"]
ws.append(cols)
for c in ws[1]: c.fill = H; c.font = HF; c.alignment = W
for r in ROWS:
    ws.append([r["v"], r["src"], r["st"], r["ent"], r["num"], r["title"], r["prod"], r["qty"], r["val"], r["due"], r["sub"], r["req"], r["fit"], r["link"]])
    for i in range(1, len(cols) + 1): ws.cell(ws.max_row, i).fill = FILL[r["vk"]]
    if isinstance(r["link"], str) and r["link"].startswith("http"): c = ws.cell(ws.max_row, 14); c.hyperlink = r["link"]; c.font = LINK
    ws.cell(ws.max_row, 9).number_format = "# ##0"
for i, w in enumerate([34, 12, 10, 30, 18, 28, 50, 30, 12, 20, 18, 50, 8, 40], 1): ws.column_dimensions[CL(i)].width = w
for row in ws.iter_rows(min_row=2):
    for c in row: c.alignment = W
ws.freeze_panes = "B2"; ws.auto_filter.ref = ws.dimensions
b = wb.create_sheet("Цены победителей")
b.append(["Продукт", "Цена / диапазон", "Детали", "Источник"])
for c in b[1]: c.fill = H; c.font = HF; c.alignment = W
for x in M.get("benchmarks", []):
    if isinstance(x, dict):
        b.append([x.get("product") or x.get("name"), txt(x.get("typical") or x.get("price") or x.get("range") or x.get("price_range")), txt({k: v for k, v in x.items() if k not in ("product", "name", "sources", "source", "links")})[:800], txt(x.get("sources") or x.get("source") or x.get("links"))[:400]])
for i, w in enumerate([28, 30, 90, 60], 1): b.column_dimensions[CL(i)].width = w
for row in b.iter_rows(min_row=2):
    for c in row: c.alignment = W
d = wb.create_sheet("Импорт и пошлины")
d.append(["Раздел", "Данные"])
for c in d[1]: c.fill = H; c.font = HF
for k in ("tariff_regime", "import_duties", "adcvd", "logistics", "landed_cost_examples"):
    v = M.get(k)
    for x in (v if isinstance(v, list) else [v]): d.append([k, txt(x)[:3000]])
d.column_dimensions["A"].width = 22; d.column_dimensions["B"].width = 170
for row in d.iter_rows(min_row=2):
    for c in row: c.alignment = W
p = wb.create_sheet("Что оформить")
for x in ["1. SAM.gov: EIN → Login.gov → UEI → полная регистрация (CAGE авто). 2–4 недели для фирмы с иностранным владельцем. Нужна для федеральных и полезна для муниципальных.",
          "2. Регистрации на порталах заказчиков: OpenGov (Collier), Bonfire (Kansas City), DemandStar/BidNet, IonWave (Choice Partners TX), PlanetBids (BACC CA — февраль 2027).",
          "3. Страхование: GL $1M/$2M, umbrella, auto $1M (HNOA), pollution liability $1–5M — ~$15–40 тыс./год. Сертификат страхования часто нужен уже при подаче.",
          "4. Бонды: заявочный ~5 %, исполнения до 100 % — кредитная линия у сюрети ($250–500 тыс., премия 1,5–4 %, личное поручительство) или программа SBA Surety Bond Guarantee.",
          "5. Опасные грузы: регистрация PHMSA ($275), обучение персонала по 49 CFR 172.704; хлор/SO2 не брать (нужен HMSP). Импорт: TSCA-сертификация; гипохлориты — FIFRA (EPA) и регистрация во Флориде.",
          "6. NSF/ANSI/CAN 60: работать только с производителями с листингом на конкретный завод, поставка в оригинальной таре/наливе с завода; иначе — своя сертификация.",
          "7. Ограничения: российское гражданство владельца не мешает обычным поставкам (малый бизнес доступен), но квоты 8(a)/WOSB/SDVOSB/HUBZone закрыты; на федеральных — Buy American/TAA (Китай и РФ не допускаются).",
          "8. Пошлины (по данным модели, проверить у брокера): Китай — MFN + Section 301 25 % + 12,5 %; Индия — MFN + 10 %; антидемпинг на активированный уголь из Китая ~$2,42/кг, хлоризоцианураты 285 %, гипохлорит кальция AD/CVD — из Китая эти товары не возить.",
          "9. Ниши для входа (модель): активированный уголь из Индии (кокосовый; где спецификация допускает), перманганат калия из Индии (при NSF 60), фосфатные ингибиторы под private label с NSF, полимеры «или эквивалент», поставки в кубах/бочках малым водоканалам. Не входить: гипохлорит, каустик, хлор, известь, CO2, сульфат Al/железо навалом."]:
    p.append([x])
p.column_dimensions["A"].width = 170
for row in p.iter_rows():
    for c in row: c.alignment = W
wb.save("../USA_ZAKUPKI_1611.xlsx")
# PDF
e = html.escape
css = ("@page{size:A4 landscape;margin:7mm}body{font-family:'DejaVu Sans',Arial;font-size:6.8px;color:#111}h1{font-size:12px;margin:0 0 2px;color:#0B2545}h2{font-size:9px;margin:5px 0 2px;color:#0B2545}"
       "table{border-collapse:collapse;width:100%}th{background:#0B2545;color:#fff;padding:2px 3px;text-align:left}td{border-bottom:1px solid #ddd;padding:1px 3px;vertical-align:top}.n{text-align:right;white-space:nowrap}"
       ".k0{background:#E6F4EA}.k1{background:#FFF4D6}.g{color:#555}a{color:#0563C1;text-decoration:none}")
cnt = {k: sum(1 for r in ROWS if r["vk"] == k) for k in range(4)}
h = [f"<!DOCTYPE html><html><head><meta charset=utf-8><style>{css}</style></head><body><h1>США: госзакупки водоочистной химии с подачей ≥ 20.11.2026 — где ISP Group может быть дешевле рынка</h1>",
     f"<div class=g>Всего {len(ROWS)}: подаём по нишевым позициям — {cnt[0]}, возможно — {cnt[1]}, только товарная химия местных заводов — {cnt[2]}, недоступно (квоты) — {cnt[3]}. Полный список, цены победителей, пошлины и что оформить — в Excel.</div>",
     "<h2>Подаём / возможно</h2><table><tr><th>Решение</th><th>Заказчик (штат)</th><th>Номер</th><th>Продукты</th><th>Оценка $/год</th><th>Подача</th><th>Требования</th></tr>"]
for r in [x for x in ROWS if x["vk"] in (0, 1)]:
    lk = f"<a href='{e(r['link'])}'>{e(str(r['num'] or ''))}</a>" if isinstance(r["link"], str) else e(str(r["num"] or ""))
    val = "—" if not isinstance(r["val"], (int, float)) else f"{r['val']:,.0f}".replace(",", " ")
    h.append(f"<tr class=k{r['vk']}><td>{e(r['v'])}</td><td><b>{e(str(r['ent'])[:60])}</b> ({e(str(r['st'])[:20])})</td><td>{lk}</td><td>{e(r['prod'][:170])}</td><td class=n>{val}</td><td>{e(str(r['due'])[:30])}</td><td class=g>{e(r['req'][:150])}</td></tr>")
h.append("</table><h2>Экономика импорта (из модели, проверить у таможенного брокера)</h2><div>Уголь кокосовый из Индии: себестоимость ~$1,10/lb при ценах победителей $1,55–2,09/lb (маржа 20–35 %). Перманганат калия из Индии: ~$1,58/lb против $2,2–5,3/lb. "
         "Китай: +Section 301 25 % +12,5 % и антидемпинг (уголь ~$2,42/кг, гипохлорит кальция, хлоризоцианураты) — по углю/гипохлориту кальция закрыт; PACl и полимеры из Китая с фрахтом $350–450/т и пошлиной ~41 % дороже цен победителей ($2,28–3,59/gal PACl; $1,15–1,80/lb полимер). "
         "Гипохлорит, каустик, хлор, известь, CO2, сульфат Al и соли железа — местные заводы, маржа 0–8 %, не входить.</div></body></html>")
hp = f"{SCRATCH}/usa.html"; open(hp, "w").write("".join(h))
subprocess.run(["/opt/pw-browsers/chromium", "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer", f"--print-to-pdf={os.path.abspath('../USA_ZAKUPKI_1611.pdf')}", "file://" + hp], check=True, capture_output=True, timeout=180)
import pypdf; print("pages", len(pypdf.PdfReader("../USA_ZAKUPKI_1611.pdf").pages), "rows", len(ROWS), cnt)
for r in [x for x in ROWS if x["vk"] == 0]: print(r["st"], "|", str(r["ent"])[:40], "|", r["dk"], "|", r["v"][:70])
