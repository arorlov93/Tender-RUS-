"""Глубокая проверка ТЗ 02.11 — закупки со сроком поставки ≤ 6 мес и маржой ≥ 25 %:
Смоленскрегионтеплоэнерго (32616488988), Россети Урал (32616531898), Чеховское ЖКХ (32616539033).
Цены — data/zakupki_0211/tz/deep/*_out.json, rosseti.json, chekhov.json (агенты, ноябрь 2026, с НДС; мы на УСН).
Заводы — m_smolA/B.json, rosseti.json, chekhov.json (контакты подтверждены в ≥ 2 источниках, иначе помечено).
→ ../DEEP_0211_ANALIZ.xlsx (внутренний: цены, маржа, риски) и ../RASSYLKA_0211_ZAVODY.xlsx (для рассылки — БЕЗ наших цен)."""
import json, os, re, openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from tidy import tidy

D = "data/zakupki_0211/tz/deep"
J = lambda f, d=None: json.load(open(f"{D}/{f}")) if os.path.exists(f"{D}/{f}") else d
thin = Side(style="thin", color="BFBFBF"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
HDR = lambda ws, r: [setattr(c, "font", Font(bold=True, color="FFFFFF")) or setattr(c, "fill", PatternFill("solid", fgColor="305496")) or setattr(c, "alignment", Alignment(wrap_text=True, vertical="top")) for c in ws[r]]
def row(ws, vals, fmt=None, fill=None):
    ws.append(vals); X = ws.max_row
    for c in ws[X]: c.border = BOX; c.alignment = Alignment(wrap_text=True, vertical="top")
    for j, f in (fmt or {}).items(): ws.cell(X, j).number_format = f
    if fill: ws.cell(X, fill[0]).fill = PatternFill("solid", fgColor=fill[1])
    return X
def widths(ws, w):
    for i, x in enumerate(w, 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = x

def profit(P, C, days):  # P — наша цена (УСН, без НДС сверху), C — себестоимость с НДС
    pre = P - C - 0.008 * P - C * 0.2 * days / 365
    return pre - 0.15 * max(pre, 0)

# ---------- Смоленск
SM = J("smolensk_full.json"); NMC_SM = {x["n"]: x for x in json.load(open("data/zakupki_0211/tz/items/Smolensk_electro.json"))}
PS = {}
for f in ("s1_out.json", "s2_out.json"):
    for x in J(f, []): PS[int(x["n"])] = x
for x in J("s4_out.json", []):
    PS[int(x["n"])] = x
if not J("s4_out.json"):
    for x in J("s4_in.json", []): PS[x["n"]] = dict(x["prev"], n=x["n"])
sm_items = []
for x in SM:
    n = x["n"]; p = PS.get(n, {}); nm = NMC_SM.get(n, {}).get("nmc_unit")
    sm_items.append(dict(n=n, name=x["name"], spec=x["spec"], unit=x["unit"], qty=x["qty"], nmc=nm, price=p.get("price"), product=p.get("product"), match=p.get("match"), source=p.get("source"), url=p.get("url"), note=p.get("unit_note")))

# ---------- Россети
RO = J("rosseti.json", {}) or {}
RQ = {int(r["n"]): r for r in RO.get("requirements", [])}
RP = {}
for p in RO.get("prices", []):
    for n in (p["n"] if isinstance(p["n"], list) else [p["n"]]):
        if p.get("price") and (int(n) not in RP or p["price"] < RP[int(n)]["price"]): RP[int(n)] = p
ro_items = []
for x in json.load(open("data/zakupki_0211/tz/items/43834_Rosseti_Ural_pipes.json")):
    r = RQ.get(x["n"], {}); p = RP.get(x["n"], {})
    ro_items.append(dict(n=x["n"], name=r.get("name_full") or x["name"], spec=r.get("reqs") or x.get("spec"), unit=x["unit"], qty=x["qty"], nmc=x["nmc_unit"], nmc_nv=r.get("nmc_unit_no_vat") or round(x["nmc_unit"] / 1.22, 2),
                         price=p.get("price"), product=(p.get("product") or "") + (f' ({p.get("maker")})' if p.get("maker") else ""), match=p.get("match"), source=p.get("source"), url=p.get("url"), note=p.get("unit_note")))

# ---------- Чехов
CH = J("chekhov.json", {}) or {}
CQ = {int(r["n"]): r for r in CH.get("requirements", [])}
CP = {}
for p in CH.get("prices", []):
    if p.get("price") and (int(p["n"]) not in CP or p["price"] < CP[int(p["n"])]["price"]): CP[int(p["n"])] = p
ch_items = []
for x in json.load(open("data/zakupki_0211/tz/items/Chekhov_spec.json")):
    r = CQ.get(x["n"], {}); p = CP.get(x["n"], {})
    ch_items.append(dict(n=x["n"], name=x["name"], spec=r.get("reqs") or "", unit=x["unit"], qty=x["qty"], nmc=r.get("nmc_unit"), price=p.get("price"),
                         product=(p.get("product") or "") + (f' ({p.get("maker")})' if p.get("maker") else ""), match=p.get("match"), source=p.get("source"), url=p.get("url"), note=p.get("price_type")))

T = [
    dict(key="Смоленск", title="Смоленскрегионтеплоэнерго — электроматериалы", eis="32616488988", dl="06.11.2026 10:00", term="30 календарных дней с даты договора, партиями по графику", nmc_total=None, items=sm_items, days=60,
         risks="Цены ЕИС — НМЦ ед. (ставка НДС не указана, принято с НДС). Эквиваленты допускаются, ±20 % размеры (кроме кабеля, наконечников, гофры, клипс, АКБ, DIN). Товар не ранее 2025 г. Доставка до 5 филиалов (Вязьма, Рославль, Сафоново, Ярцево, Смоленск) — за счёт поставщика."),
    dict(key="Россети", title="Россети Урал — трубы для прокладки КЛ (43834)", eis="32616531898", dl="10.11.2026 09:00", term=RO.get("delivery") or "30 календарных дней с даты договора (в НМЦ — разбивка по кварталам потребности 2027)", nmc_total=None, items=ro_items, days=45,
         risks="Все позиции — аттестация в ПАО «Россети» (№ заключения в заявке; непредставление в 90 дн. — нарушение). Обеспечение исполнения 5 % НМЦ. Оплата 7 раб. дн. Доставка Пермский край, Свердловская, Челябинская обл. " + (RO.get("regime") or "")),
    dict(key="Чехов", title="Чеховское ЖКХ — спецодежда, спецобувь", eis="32616539033", dl="11.11.2026 11:00", term="по заявкам, 5 раб. дней с заявки, до 31.12.2026", nmc_total=3077114.99, items=ch_items, days=40,
         risks="НМЦД 3 077 114,99 с НДС 22 %. СИЗ не ранее 2026 г., ТР ТС 019/2011. Оплата 7 раб. дн. Поставка мелкими заявками — логистика по Чехову."),
]

# ---------- 1) Анализ (внутренний)
wb = openpyxl.Workbook(); sv = wb.active; sv.title = "Итог"
sv.append(["Глубокая проверка: срок поставки ≤ 6 мес, маржа ≥ 25 %. Рынок — закупка с НДС (мы на УСН). Прибыль: −0,8 % площадка, −20 % годовых на оборот, −15 % УСН."]); sv["A1"].font = Font(bold=True, size=12)
sv.append(["Закупка", "№ ЕИС", "Подача", "Срок поставки", "НМЦ, ₽", "Позиций", "С подтверждённой ценой", "Покрытие НМЦ", "НМЦ покрытой части", "Рынок покрытой части", "Маржа", "Маржа при −10 %", "Чистая прибыль при −10 %, ₽ (на покрытую часть)", "Вывод", "Риски и условия"]); HDR(sv, 2)
for t in T:
    L = t["items"]; nm_all = sum((x["nmc"] or 0) * (x["qty"] or 0) for x in L); tot = t["nmc_total"] or nm_all
    cov = [x for x in L if x["price"]]
    if all(x["nmc"] for x in L):
        nm = sum(x["nmc"] * x["qty"] for x in cov); share = nm / nm_all if nm_all else 0
    else:  # НМЦ ед. неизвестна — только при полном покрытии
        nm = tot if len(cov) == len(L) else None; share = len(cov) / len(L) if L else 0
    mk = sum(x["price"] * x["qty"] for x in cov)
    mg = (nm - mk) / nm if nm else None; mg10 = (nm * 0.9 - mk) / (nm * 0.9) if nm else None
    pr = profit(nm * 0.9, mk, t["days"]) if nm else None
    t.update(mg=mg, mg10=mg10, tot=tot)
    v = "нет данных" if mg10 is None else "ИДЁМ" if mg10 >= 0.25 else "ИДЁМ, тонко" if mg10 >= 0.15 else "НЕ ИДЁМ"
    row(sv, [t["title"], t["eis"], t["dl"], t["term"], tot, len(L), len(cov), share, nm, mk, mg, mg10, pr, v, t["risks"]], {5: "#,##0", 9: "#,##0", 10: "#,##0", 13: "#,##0", 8: "0%", 11: "0%", 12: "0%"},
        (14, "C6EFCE" if v == "ИДЁМ" else "FFF2CC" if v.startswith("ИДЁМ") else "F8CBAD" if v.startswith("НЕ") else "D9D9D9"))
sv.append([]); sv.append(["Отсеяны: Газпром теплоэнерго МО 5044 (32616526616) и Киров 5036 (32616475914) — поставка по заявкам весь 2027 г. (> 6 мес). Остальные ТЗ партии — маржа < 25 % (см. TZ_0211_MARGIN.xlsx)."])
widths(sv, [36, 13, 12, 26, 13, 8, 10, 9, 13, 13, 8, 9, 14, 12, 70])
for t in T:
    ws = wb.create_sheet(f'{t["key"]} — цены')
    ws.append(["№", "Позиция ТЗ", "Требования ТЗ", "Ед.", "Кол-во", "НМЦ ед., ₽", "Рынок ед., ₽ (с НДС)", "Маржа ед.", "Сумма НМЦ", "Сумма рынок", "Товар (модель, производитель)", "Соответствие ТЗ", "Примечание к цене", "Источник", "Ссылка"]); HDR(ws, 1)
    for x in t["items"]:
        m = (x["nmc"] - x["price"]) / x["nmc"] if x["nmc"] and x["price"] else None
        X = row(ws, [x["n"], x["name"], (x["spec"] or "")[:600], x["unit"], x["qty"], x["nmc"], x["price"], m, (x["nmc"] or 0) * x["qty"] or None, (x["price"] or 0) * x["qty"] or None, x["product"], x["match"], x["note"], x["source"], x["url"]],
                {6: "#,##0.00", 7: "#,##0.00", 8: "0%", 9: "#,##0", 10: "#,##0"},
                (8, "C6EFCE" if (m or 0) >= 0.25 else "FFF2CC" if (m or 0) >= 0 else "F8CBAD") if m is not None else (7, "D9D9D9"))
        if x["url"] and str(x["url"]).startswith("http"): ws.cell(X, 15).hyperlink = x["url"]
    widths(ws, [5, 36, 50, 6, 7, 11, 11, 7, 11, 11, 30, 26, 18, 14, 30]); ws.freeze_panes = "C2"; ws.auto_filter.ref = f"A1:O{ws.max_row}"
tidy(wb, row_h=30); wb.save("../DEEP_0211_ANALIZ.xlsx")

# ---------- 2) Рассылка заводам (без цен)
CATS_SM = {}
for f in ("m_smolA.json", "m_smolB.json"):
    for k, v in (J(f, {}) or {}).items(): CATS_SM[k] = v
def rng(s):
    out = set()
    for a in re.findall(r"\d+(?:\s*[–-]\s*\d+)?", str(s)):
        b = re.split(r"\s*[–-]\s*", a); out |= set(range(int(b[0]), int(b[-1]) + 1))
    return out
ITEMS = {t["key"]: {x["n"]: x for x in t["items"]} for t in T}
blocks = {"Смоленск": [(v.get("title", k), sorted(rng(v.get("positions", ""))), v.get("makers", [])) for k, v in sorted(CATS_SM.items(), key=lambda kv: (kv[0][0], int(kv[0][1:])))],
          "Россети": [("Трубы двустенные гофрированные ПНД с зондом и труба гладкая SN12", list(range(1, 19)), RO.get("makers", []))],
          "Чехов": [("Спецодежда и спецобувь", list(range(1, 13)), CH.get("makers", []))]}
wr = openpyxl.Workbook(); s0 = wr.active; s0.title = "Как пользоваться"
for r in [["Рассылка запросов КП заводам — ГИДРОПАС, ноябрь 2026. Цен в файле нет."], [],
          ["Лист «<закупка> — кому что»: по каждой категории — позиции ТЗ для запроса и 6–7 заводов с контактами."],
          ["Лист «<закупка> — позиции ТЗ»: полные требования ТЗ — прикладывать к запросу."],
          ["Контакты: «да» — телефон/e-mail подтверждены в 2 независимых источниках (ссылки в последней колонке); «нет» — перепроверить перед отправкой."],
          ["Срок ответа просить: Смоленск — до 05.11 (подача 06.11 10:00), Россети — до 08.11, Чехов — до 09.11."]]:
    s0.append(r)
s0["A1"].font = Font(bold=True, size=12); s0.column_dimensions["A"].width = 120
for t in T:
    k = t["key"]; ws = wr.create_sheet(f"{k} — кому что")
    ws.append([f'{t["title"]} — ЕИС {t["eis"]}, подача {t["dl"]}, поставка: {t["term"]}']); ws["A1"].font = Font(bold=True, size=12)
    ws.append(["Категория", "Позиции ТЗ для запроса (№ — наименование — кол-во)", "Завод / поставщик", "ИНН", "Город", "Сайт", "Телефон", "E-mail", "Что выпускает", "Примечание", "Подтверждено в 2 источниках", "Источники контактов"]); HDR(ws, 2)
    for title, ns, makers in blocks[k]:
        pos = "\n".join(f'{n} — {ITEMS[k][n]["name"][:70]} — {ITEMS[k][n]["qty"]:g} {ITEMS[k][n]["unit"] or ""}' for n in ns if n in ITEMS[k])
        if len(pos) > 3000: pos = pos[:3000] + "\n… полный перечень — лист «позиции ТЗ»"
        if not makers: row(ws, [title, pos, "— заводы не найдены —"]); continue
        for i, m in enumerate(makers):
            ra = m.get("rosseti_attest") or {}
            note = (m.get("note") or "") + (f' Аттестация Россети: {"есть" if ra.get("has") else "нет" if ra.get("has") is False else "не выяснено"} {ra.get("number") or ""}' if k == "Россети" else "")
            pp = m.get("positions") if k != "Смоленск" and m.get("positions") else None
            row(ws, [title if i == 0 else "", (pos if i == 0 else "") if not pp else str(pp), m.get("name"), m.get("inn"), m.get("city"), m.get("site"), m.get("phone"), m.get("email"), m.get("products"), note.strip(),
                     "да" if m.get("verified") else "нет", "\n".join(m.get("sources") or [])], fill=(11, "C6EFCE" if m.get("verified") else "FFF2CC"))
    widths(ws, [24, 50, 28, 12, 12, 20, 18, 24, 30, 30, 9, 34]); ws.freeze_panes = "C3"
    w2 = wr.create_sheet(f"{k} — позиции ТЗ"); w2.append(["№", "Наименование", "Требования ТЗ (дословно)", "Ед.", "Кол-во"]); HDR(w2, 1)
    for x in t["items"]: row(w2, [x["n"], x["name"], x["spec"], x["unit"], x["qty"]])
    widths(w2, [5, 40, 100, 6, 8])
tidy(wr, row_h=30); wr.save("../RASSYLKA_0211_ZAVODY.xlsx")
for t in T: print(t["key"], "маржа", t["mg"] and round(t["mg"], 3), "−10%", t["mg10"] and round(t["mg10"], 3))
print({k: sum(len(b[2]) for b in v) for k, v in blocks.items()})
