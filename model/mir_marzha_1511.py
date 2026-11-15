"""Детальная маржа по мировым тендерам (15.11.2026).
Вход: data/tenders_1511/spec_*.json (спецификации от агентов), cost_benchmarks.json (FOB, фрахт, пошлины), mir_ru.json, fx_usd_1511.json, rows_1511.json.
Выход: data/tenders_1511/mir_marzha.json + ../ZAKUPKI_MIR_1511.xlsx (лист «Маржа детально») + ../ZAKUPKI_MIR_MARZHA_1511.pdf"""
import json, re, glob, os, html, subprocess, statistics
from isp_letterhead import SCRATCH
T = "data/tenders_1511/"
CB = json.load(open(T + "cost_benchmarks.json"))
FX = json.load(open(T + "fx_usd_1511.json"))["rates"]; FX["USD"] = 1.0; FX.setdefault("EURO", FX["EUR"])
MR = {x["id"]: x for x in json.load(open(T + "mir_ru.json"))}
R = json.load(open(f"{SCRATCH}/rows_1511.json")); MIR = [r for r in R if r["who"] == "ISP"]
P = {p["key"]: p for p in CB["products"]}
def fob(k): return (P.get(k) or {}).get("fob_typical")

# ---------- допущения (явно помечены)
A = dict(payload_t=20.0, inland_eu=40.0, inland_far=120.0, other_pct=0.03, dg=200.0)
# ---------- продукт по тексту → ключ себестоимости, HS-группа, класс опасности
RULES = [
    (r"pbtc|фбтк", "pbtca_50", "2931.49.80", 1), (r"hedp|оэдф", "hedp_60", "2931.49", 1), (r"atmp|нтф", "atmp_50", "2931.49", 1),
    (r"antiscal|антискал|scale inhib.*ro|ro.*antiscal|ингибитор накип|mts?i ?410|hypersperse", "ro_antiscalant", "3824.99", 0),
    (r"cooling|градирн|оборотн|corrosion.*scale|scale.*corrosion|ингибитор коррозии и накип|dispersant|диспергатор", "cooling_scale_corr", "3824.99", 0),
    (r"demuls|деэмульг", "demulsifier_oil_sol", "3824.99", 0), (r"h2s\b|сероводор|triazin|триазин|scaveng.*sulf", "h2s_triazine_40", "3824.99", 0),
    (r"pour point|депрессор|ppd|paraffin|парафин", "ppd_paraffin", "3824.99", 0),
    (r"corrosion inhib|ингибитор коррозии|селерин|imidazol", "ci_imidazoline", "3824.99", 0),
    (r"oxygen scav|поглотитель кислор|carbohydraz|карбогидраз|deha", "carbohydrazide", "3824.99", 0),
    (r"bisulf|бисульф|metabisulf", "na_bisulfite", "2832", 0), (r"morpholin|морфолин|cyclohexylam|циклогексил|amine|амин", "morpholine", "2921", 0),
    (r"trisodium phos|тринатрийфосф|tsp\b", "tsp", "2835", 0), (r"isothiaz|изотиаз|biocid|биоцид|algicid|альгицид|glutar", "isothiazolinone_1_5", "3808", 0),
    (r"membrane clean|очистител|cip", "citric_acid", "3824.99", 0),
    (r"poly ?dadmac|полидадмак|organic coag|органическ.*коагул|polyamin|полиамин", "polydadmac_40", "3906.90", 0),
    (r"emuls.*(polyelect|polimer|pam|flocc|флокул|полиэлект|полиакрил)|(polyelect|pam|flocc|флокул|полиэлект|полиакрил).*emuls|эмульс", "pam_cationic_emulsion", "3906.90", 0),
    (r"anion|анион", "pam_anionic_listing", "3906.90", 0),
    (r"polyacryl|полиакрил|polyelect|полиэлект|flocc|флокул|\bpam\b|\bпаа\b|polimer|polymer|полимер", "pam_cationic", "3906.90", 0),
    (r"ferric chlor|хлорн\w* желез|fecl3|chlorure ferrique|cloruro f[eé]rrico|chlorek żelaza|demir.*klor", "fecl3_liquid_40", "2827.39", 1),
    (r"ferric sulph|ferric sulf|сульфат желез|сернокисл\w* желез|siarczan żelaza|iron.*sulph|pix|fe2\(so4\)", "fe2so4_3_solid", "2833.29", 0),
    (r"polyferric|pfs", "pfs_polyferric", "2833.29", 0), (r"ferrous|железн\w* купорос|feso4", "feso4_7h2o", "2833.29", 0),
    (r"pac\b|poly ?alumin|полиалюм|полиоксихлор|пах\b|поха|polichlorek glinu|polialüminyum|aluminium chlorohydr|alcl3|хлорид алюмин", "PAC", "2827.49", 0),
    (r"alum|сульфат алюмин|sulfate d'alumin|sulfato de alumin|siarczan glinu|alüminyum sülfat|aluminium sulphate|aluminum sulfate", "ALUM", "2833.22", 0),
    (r"coconut|кокос", "ac_gac_coconut", "3802.10", 0), (r"powder.*carbon|порошк\w* уг|pac carbon|\bpac\b.*carbon", "ac_powder", "3802.10", 0),
    (r"carbon|уголь|углем|charbon|carbone|carbón|aktivkohle", "ac_gac_coal", "3802.10", 0),
    (r"activated alumin|оксид алюмин|alumina", "act_alumina_3_5", "2818.20", 0), (r"molecular sieve|молекуляр|zeolit|цеолит", "mol_sieve", "3802.90", 0),
    (r"silica gel|силикагел", "silica_gel", "2811", 0), (r"anion.*resin|анионит|sba|wba", "IEX_A", "3914.00", 0),
    (r"resin|смол|cation|катионит|sac\b", "iex_001x7_t", "3914.00", 0),
    (r"calcium hypo|гипохлорит кальц|hth", "ca_hypo_70", "2828.10", 1), (r"sodium hypo|гипохлорит натр|naocl|eau de javel", "naocl_12", "2828.90", 1),
    (r"tcca|трихлор", "tcca_90", "2933", 0), (r"soda ash|sodium carbon|кальцинир|карбонат натр", "soda_ash", "2836.20", 0),
    (r"lime|извест|chaux|cal hidrat", "hydrated_lime", "2522", 0), (r"caustic|naoh|едк|каустик|гидроксид натр", "naoh_flakes", "2815.11", 1),
    (r"citric|лимонн", "citric_acid", "2918", 0), (r"sulfuric|sulphuric|серн\w* кисл|h2so4", "h2so4_98", "2807", 1), (r"hydrochlor|соляная|hcl", "hcl_33", "2806", 1),
]
def classify(text):
    t = text.lower()
    for rx, key, hs, dg in RULES:
        if re.search(rx, t): return key, hs, dg
    return None, None, 0
def al2o3(text):
    m = re.search(r"(\d{1,2}(?:[.,]\d)?)\s*%?\s*(?:±\s*\d+\s*%?)?\s*(?:al2o3|al₂o3|al2 o3)|(?:al2o3|al₂o3)\D{0,25}?(\d{1,2}(?:[.,]\d)?)\s*%", text.lower())
    if not m: return None
    v = float((m.group(1) or m.group(2)).replace(",", ".")); return v if 1 <= v <= 32 else None
def unit_cost(key, text):
    """FOB USD за тонну ТОВАРА в форме, которую просит заказчик (+ пояснение)"""
    t = text.lower()
    if key == "PAC":
        liquid = bool(re.search(r"liquid|жидк|solution|раствор|ciek|sıvı|liquide|líquid|\b1[0-8](?:[.,]\d)?\s*%", t)) and not re.search(r"powder|порош|proszek|toz", t)
        a = al2o3(text)
        if liquid:
            a = a or 10.0
            return fob("pac_powder_tech") * a / 30.0, f"жидкий PAC {a:g}% Al2O3 → считаем из порошка 30% ($295/т) по Al2O3; ⚠ возить воду из Китая дорого — реально нужен местный растворитель/партнёр"
        return (fob("pac_powder_nsf") if re.search(r"nsf|drink|питьев|potable|pitn", t) else fob("pac_powder_tech")), "PAC порошок — наша КП Aierfuke"
    if key == "ALUM":
        liquid = bool(re.search(r"liquid|жидк|solution|раствор|ciek|sıvı|liquide|líquid", t))
        if liquid:
            a = al2o3(text) or 8.0
            return fob("alum_sulph_ironfree") * a / 17.0, f"жидкий сульфат алюминия {a:g}% Al2O3 → пересчёт с твёрдого 17%; ⚠ возить раствор невыгодно"
        return (fob("alum_sulph_ironfree") if re.search(r"iron.?free|безжелез|non.?ferric|low iron", t) else fob("alum_sulph_index")), "сульфат алюминия твёрдый — индекс/листинги FOB Китай"
    if key == "IEX_A": return fob("iex_201x7") / 0.70, "анионит 201x7: $2 090/м³ ÷ 0,70 т/м³ (Suqing)"
    return fob(key), (P.get(key) or {}).get("name_ru") or key

# ---------- страна → порт (ключ фрахта), доставка, страна пошлины
CTRY = {
 "Польша": ("Gdansk", "EU", "eu"), "Германия": ("Hamburg", "EU", "eu"), "Италия": ("Genoa/La Spezia", "EU", "eu"), "Испания": ("Valencia/Barcelona", "EU", "eu"),
 "Португалия": ("Valencia/Barcelona", "EU", "eu"), "Франция": ("Genoa/La Spezia", "EU", "eu"), "Швеция": ("Gothenburg", "EU", "eu"), "Дания": ("Aarhus", "EU", "eu"),
 "Финляндия": ("Gothenburg", "EU", "eu"), "Норвегия": ("Oslo", "EU", "eu"), "Венгрия": ("Constanta", "EU", "far"), "Румыния": ("Constanta", "EU", "eu"), "Болгария": ("Constanta", "EU", "eu"),
 "Словения": ("Genoa/La Spezia", "EU", "eu"), "Кипр": ("Piraeus", "EU", "eu"), "Швейцария": ("Genoa/La Spezia", "EU", "far"), "Сербия": ("Constanta", None, "far"), "Косово": ("Constanta", None, "far"),
 "Великобритания": ("Felixstowe", "UK", "eu"), "Украина": ("Constanta", "Ukraine", "far"), "Турция": ("Mersin/Istanbul", "Turkey", "eu"),
 "США": ("Houston", "USA", "eu"), "Канада": ("Montreal", "Canada", "eu"), "Бразилия": ("Santos", "Brazil", "eu"),
 "Саудовская Аравия": ("Dammam/Jeddah", "Saudi", "eu"), "Кувейт": ("Shuwaikh (Kuwait)", "Kuwait", "eu"), "Бахрейн": ("Bahrain (Khalifa Bin Salman)", "Bahrain", "eu"),
 "Оман": ("Sohar", "Oman", "eu"), "Иордания": ("Jebel Ali", None, "far"), "Египет": ("Piraeus", None, "eu"), "Марокко": ("Casablanca", "Morocco", "eu"), "Тунис": ("Rades (Tunis)", "Tunisia", "eu"),
 "Кения": ("Mombasa", "Kenya", "eu"), "Уганда": ("Mombasa", "Uganda", "far"), "Руанда": ("Dar es Salaam", None, "far"), "Малави": ("Dar es Salaam", None, "far"), "Эфиопия": ("Jebel Ali", None, "far"),
 "ЮАР": ("Durban", "South Africa", "eu"), "Зимбабве": ("Durban", None, "far"), "Маврикий": ("Durban", None, "eu"), "Мадагаскар": ("Durban", None, "eu"), "Камерун": ("Douala", "Cameroon", "eu"), "Либерия": ("Douala", None, "eu"),
 "Бангладеш": ("Chittagong", "Bangladesh", "eu"), "Пакистан": ("Karachi", "Pakistan", "eu"), "Вьетнам": ("Ho Chi Minh", "Vietnam", "eu"), "Филиппины": ("Manila", "Philippines", "eu"),
 "Таиланд": ("Laem Chabang", "Thailand", "eu"), "Австралия": ("Brisbane", "Australia", "eu"), "Индия": ("INDIA", "India", "eu"),
 "Казахстан": ("Almaty (KZ)", "Kazakhstan", "far"), "Узбекистан": ("UZ_TRUCK", "Uzbekistan", "eu"), "Туркменистан": ("UZ_TRUCK", None, "eu"), "Азербайджан": ("Poti/Batumi", None, "far"),
}
def freight_t(port, dg):
    if port == "INDIA": return 175.0, "CIF Nhava Sheva − FOB по КП Kairui (+$175/т)"
    if port == "UZ_TRUCK": return 17760 / 22, "авто Китай→Каракалпакстан $17 760/22 т (Saada Log, 13.11)"
    rows = [f for f in CB["freight"] if f["to_port"] == port and f.get("usd_typical") and "Nhava" not in f["from"]]
    r20 = next((f for f in rows if f["container"] == "20'"), None); r40 = next((f for f in rows if f["container"] == "40'"), None)
    if r20: v, b = r20["usd_typical"] / A["payload_t"], f"20' ${r20['usd_typical']:,.0f} / {A['payload_t']:g} т ({port})"
    elif r40: v, b = r40["usd_typical"] / 25.0, f"40' ${r40['usd_typical']:,.0f} / 25 т ({port})"
    else: return None, "нет ставки"
    if dg: v += A["dg"] / A["payload_t"]; b += " + DG класс 8"
    return v, b
def duty(country_key, hs, key):
    if not country_key: return 0.05, "5% — ДОПУЩЕНИЕ (ставки нет в базе)"
    ck = {"EU": "EU (TARIC", "UK": "UK", "Saudi": "Saudi", "Kuwait": "Kuwait", "Bahrain": "Bahrain", "Oman": "Oman", "Kenya": "Kenya", "Uganda": "Uganda", "South Africa": "South Africa",
          "Kazakhstan": "Kazakhstan", "Brazil": "Brazil", "Cameroon": "Cameroon"}.get(country_key, country_key)
    cand = [d for d in CB["duties"] if d["country"].startswith(ck) and isinstance(d.get("rate_pct"), (int, float)) and str(d["hs"])[:4] == hs[:4]]
    if key == "pbtca_50" and country_key == "EU": return 1.922 + 0.065, "6,5% + антидемпинг ЕС на PBTC из КНР 192,2% (Reg 2026/2088)"
    if not cand: return 0.05, f"5% — ДОПУЩЕНИЕ (нет ставки {country_key} {hs})"
    d = cand[0]; extra = 0.125 if country_key == "USA" else 0.0
    return d["rate_pct"] / 100 + extra, f"{d['rate_pct']}%" + (" + 12,5% Section 301 (КНР)" if extra else "") + f" ({d['country']} {d['hs']})"

# ---------- выручка
UNIT = [(r"/\s*(t|т|mt|ton(?!ne)?s?\b|tonne|mg\b|ton\b)", 1.0), (r"/\s*(kg|кг)", 1000.0), (r"short ton", 1 / 1.10231), (r"/\s*(solution )?lb", 2204.62), (r"dry ton", 1 / 1.10231)]
def bench_usd_t(b):
    cp = str(b.get("currency_per") or ""); what = str(b.get("what") or "")
    if b.get("unit_price") is None or re.search(r"fob|market|рыноч|cif|dry ton|dry\b", (cp + what).lower()) or re.search(r"total|contract|lot|m3|gal|per ", cp.lower()): return None
    cur = next((c for c in sorted(FX, key=len, reverse=True) if re.match(rf"\s*{c}\b", cp.upper())), None)
    if not cur: return None
    mult = next((m for rx, m in UNIT if re.search(rx, cp.lower())), None)
    if mult is None: return None
    v = b["unit_price"] * mult / FX[cur]
    return v if 30 < v < 30000 else None
def est_usd(e):
    if not e or not e.get("amount") or not e.get("currency"): return None
    c = str(e["currency"]).upper()[:3]
    return e["amount"] / FX[c] if c in FX else None

# ---------- сборка
SPEC = {}
for f in sorted(glob.glob(T + "spec_*.json")):
    for t in json.load(open(f)).get("tenders", []):
        SPEC[t.get("id") if isinstance(t.get("id"), int) else f"{f[len(T)+5:-5]}:{t.get('n')}"] = t
CLOSED = r"only (class|local)|class.?1|class.?2|make in india|\bmii\b|reserved|только (мест|местн|казахстан|отечеств|турец|маврик|тунис|омск|для)|yerli|sadece|national competitive|\bnct\b|\bncb\b|csd|b-bbee|local(ly)? produc|местного производства|nalco|single tender|sme \(local\)|квалификационн\w* спис|prequalif\w* list|registered in (tunisia|kuwait)|kuwaiti|saudi cr|bahraini cr"
out = []
for key, t in SPEC.items():
    m = MR.get(key, {}) if isinstance(key, int) else {}
    row = MIR[key] if isinstance(key, int) and key < len(MIR) else {}
    country = m.get("country_ru") or ("Камерун" if "CAMWATER" in str(t.get("n")) else "Мадагаскар" if "JIRAMA" in str(t.get("n")) else "Узбекистан" if "T-47" in str(t.get("n")) else "Оман" if "1131" in str(t.get("n")) else None)
    port, dck, inl = CTRY.get(country, (None, None, "far"))
    items = t.get("items") or []
    text_all = " ".join([str(m.get("ru") or ""), str(row.get("subj") or "")] + [f"{i.get('product','')} {i.get('spec','')}" for i in items])
    lines, cost_total, q_total, notes = [], 0.0, 0.0, []
    for it in items:
        q = it.get("qty_t")
        if not isinstance(q, (int, float)) or q <= 0: continue
        txt = f"{it.get('product','')} {it.get('spec','')}"
        k, hs, dg = classify(txt)
        if not k: lines.append(dict(item=it.get("product"), qty_t=q, note="не наш профиль / не классифицирован")); continue
        uc, why = unit_cost(k, txt)
        if not uc: lines.append(dict(item=it.get("product"), qty_t=q, note=f"нет FOB по {k}")); continue
        fr, frb = freight_t(port, dg) if port else (None, "нет порта")
        du, dub = duty(dck, hs, k)
        inland = A["inland_eu"] if inl == "eu" else A["inland_far"]
        landed = (uc + (fr or 0)) * (1 + du) + inland
        landed *= (1 + A["other_pct"])
        lines.append(dict(item=it.get("product"), qty_t=q, key=k, fob=round(uc), freight=round(fr or 0), duty_pct=round(du * 100, 1), inland=inland, landed_t=round(landed), why=why, fr_basis=frb, duty_basis=dub))
        cost_total += landed * q; q_total += q
    if not lines or q_total == 0:
        q = t.get("total_qty_t") if isinstance(t.get("total_qty_t"), (int, float)) else m.get("qty_t")
        k, hs, dg = classify(text_all)
        if q and k:
            uc, why = unit_cost(k, text_all)
            if uc:
                fr, frb = freight_t(port, dg) if port else (None, "нет порта")
                du, dub = duty(dck, hs, k); inland = A["inland_eu"] if inl == "eu" else A["inland_far"]
                landed = ((uc + (fr or 0)) * (1 + du) + inland) * (1 + A["other_pct"])
                lines = [dict(item="(весь объём)", qty_t=q, key=k, fob=round(uc), freight=round(fr or 0), duty_pct=round(du * 100, 1), inland=inland, landed_t=round(landed), why=why, fr_basis=frb, duty_basis=dub)]
                cost_total, q_total = landed * q, q
    ev = est_usd(t.get("estimated_value")) or (m.get("value_amount") / FX[m["value_currency"]] if m.get("value_amount") and m.get("value_currency") in FX else None)
    fam = lambda k: {"pam_cationic_emulsion": "PAM", "pam_cationic": "PAM", "pam_anionic_listing": "PAM", "fecl3_liquid_40": "FE", "fe2so4_3_solid": "FE", "pfs_polyferric": "FE"}.get(k, k)
    BB = [(classify(str(b.get("what")))[0], bench_usd_t(b)) for b in (t.get("benchmark_prices") or [])]
    BB = [(k, v) for k, v in BB if v]
    rev_lines, matched = 0.0, 0
    for l in lines:
        if not l.get("key"): continue
        vs = [v for k, v in BB if k and fam(k) == fam(l["key"])] or ([v for k, v in BB] if len({fam(x["key"]) for x in lines if x.get("key")}) == 1 and len(BB) and all(not k or fam(k) == fam(l["key"]) for k, v in BB) else [])
        if vs: l["bench_t"] = round(statistics.median(vs)); rev_lines += l["bench_t"] * l["qty_t"]; matched += 1
    costed = [l for l in lines if l.get("key")]
    bp = (rev_lines / sum(l["qty_t"] for l in costed)) if costed and matched == len(costed) else None
    incomplete = any(not l.get("key") for l in lines) or any((not isinstance(i.get("qty_t"), (int, float)) or not i.get("qty_t")) for i in items if classify(f"{i.get('product','')} {i.get('spec','')}")[0]) if items else False
    warn = ""
    if ev and q_total and not incomplete:
        rev, rev_src = ev, "опубликованная стоимость тендера"
        ratio = (ev / q_total) / (cost_total / q_total)
        if ratio > 4: warn = f"⚠ стоимость/т в {ratio:.0f}× выше нашей себестоимости — вероятно, стоимость за весь срок/с опционами или включает другие позиции; проверить"
        if ratio < 1: warn = "⚠ стоимость ниже себестоимости — не проходим"
    elif bp and q_total: rev, rev_src = bp * q_total, f"цена прошлых контрактов/бюджета ≈${bp:,.0f}/т × объём"
    elif ev and q_total and incomplete: rev, rev_src = ev, "стоимость всего тендера; часть позиций без объёма/вне профиля — маржу не считаем"
    elif ev: rev, rev_src = ev, "опубликованная стоимость (объём неизвестен — себестоимость не посчитана)"
    else: rev, rev_src = None, "нет данных о цене"
    margin = (rev - cost_total) if (rev and q_total and cost_total and "маржу не считаем" not in rev_src) else None
    elig = " ".join(str(t.get(k) or "") for k in ("eligibility_foreign", "feasibility_ru")).lower()
    closed = bool(re.search(CLOSED, elig))
    acc = "Закрыт для ISP напрямую (только местные/бренд/квал. список)" if closed else (m.get("access_ru") or "Уточнить")
    out.append(dict(id=key, n=t.get("n"), country=country, buyer=row.get("cust") or "", ru=m.get("ru") or row.get("subj") or "", dl=row.get("dl"), link=row.get("link"),
                    docs=bool(t.get("docs_found")), qty_t=round(q_total, 1) if q_total else (t.get("total_qty_t") or m.get("qty_t")), lines=lines,
                    cost_usd=round(cost_total) if cost_total else None, rev_usd=round(rev) if rev else None, rev_src=rev_src, bench_t=round(bp) if bp else None,
                    margin_usd=round(margin) if margin is not None else None, margin_pct=round(margin / rev * 100, 1) if margin is not None and rev else None,
                    breakeven_t=round(cost_total / q_total) if q_total and cost_total else None, access=acc,
                    period=t.get("contract_period"), delivery=t.get("delivery_terms"), certs=t.get("certificates"), security=t.get("bid_security"),
                    feas=t.get("feasibility_ru"), risks=t.get("risks_ru"), warn=warn, doc_urls=(t.get("doc_urls") or [])[:3]))
json.dump(dict(assumptions=A, rows=out), open(T + "mir_marzha.json", "w"), ensure_ascii=False, indent=1, default=str)
calc = [o for o in out if o["margin_usd"] is not None]
print(len(out), "tenders;", len(calc), "with margin;", sum(1 for o in out if o["cost_usd"]), "with cost")
for o in sorted(calc, key=lambda o: -o["margin_usd"])[:25]:
    print(o["id"], o["country"], o["qty_t"], o["rev_usd"], o["cost_usd"], o["margin_usd"], o["margin_pct"], o["access"][:12], "|", o["rev_src"][:30], "|", [l.get("key") for l in o["lines"]][:3])
