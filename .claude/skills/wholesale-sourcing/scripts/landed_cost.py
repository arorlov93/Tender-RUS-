#!/usr/bin/env python3
"""Себестоимость единицы товара в Москве: поставщик в ЕС → (растаможка в Беларуси через импортёра | транзит и растаможка в РФ) → Москва.
Считает за единицу (банку/бутылку) и на весь заказ: без НДС, с одним НДС (НДС РБ возмещён), с двумя НДС (НДС РБ не возмещён).
Использование: python3 landed_cost.py config.json [--out ../ИМЯ]   → ИМЯ.xlsx, ИМЯ.html (+ ИМЯ.pdf, если найден chromium)
Формат config.json — см. examples/redbull_vievita.json.
"""
import sys, json, os, html, subprocess, shutil

cfg = json.load(open(sys.argv[1]))
out = sys.argv[sys.argv.index("--out") + 1] if "--out" in sys.argv else os.path.splitext(sys.argv[1])[0]
R = cfg["rates"]; EUR, BYN = cfg["fx"]["EUR"], cfg["fx"]["BYN"]
units = sum(i["cases"] * i["units_per_case"] for i in cfg["items"])
by = cfg.get("route", "BY") == "BY"


def calc(it):
    vol = it["litres_per_unit"]
    p = it["price_eur_per_case"] / it["units_per_case"] * EUR
    bank = p * R["bank_agent"]
    fr1 = cfg["freight"]["to_eaeu_eur"] * EUR / units
    fr2 = cfg["freight"]["to_moscow_rub"] / units
    cv = p + fr1
    duty = max(cv * R["duty"], R.get("duty_min_eur_per_l", 0) * vol * EUR)
    by_exc = R["by_excise_byn_per_l"] * vol * BYN if (by and it.get("by_excise", True)) else 0
    fees = (cfg["fixed"]["by_rub"] if by else cfg["fixed"]["rf_rub"]) / units
    imp = (cv + duty) * R["by_importer"] if by else 0
    mark = R["sticker"] + R["marking_code"] + R["warehouse"]
    rf_exc = R["rf_excise_rub_per_l"] * vol if it.get("rf_excise", False) else 0
    net = p + bank + fr1 + fr2 + duty + by_exc + fees + imp + mark + rf_exc
    by_vat = (cv + duty + by_exc) * R["by_vat"] if by else 0
    base = cv + duty + by_exc + fees + imp + R["sticker"] + R["marking_code"]
    v1 = (base + rf_exc) * R["rf_vat"]
    v2 = (base + by_vat + rf_exc) * R["rf_vat"] if by else v1
    return dict(price=p, bank=bank, freight=fr1 + fr2, duty=duty, by_exc=by_exc, fees=fees + imp, mark=mark, rf_exc=rf_exc,
                net=net, rf_vat1=v1, one=net + v1, by_vat=by_vat, rf_vat2=v2, two=net + by_vat + v2)


ROWS = [("Цена поставщика за ед.", "price"), ("Банк / платёжный агент", "bank"), ("Фрахт до ЕАЭС и до Москвы", "freight"), ("Пошлина ЕАЭС", "duty"),
        ("Акциз РБ", "by_exc"), ("Брокер, сборы, СВХ, декларация, импортёр", "fees"), ("Стикер, «Честный знак», склад", "mark"), ("Акциз РФ", "rf_exc"),
        ("ИТОГО без НДС", "net"), ("НДС РФ (один НДС)", "rf_vat1"), ("ИТОГО с одним НДС", "one"), ("НДС РБ не возмещён", "by_vat"),
        ("НДС РФ (с учётом НДС РБ)", "rf_vat2"), ("ИТОГО с двумя НДС", "two")]
res = [(it, calc(it)) for it in cfg["items"]]
tot = lambda k: sum(r[k] * it["cases"] * it["units_per_case"] for it, r in res)

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "₽ за единицу"
    ws.append(["₽ за единицу"] + [f"{it['name']} (€{it['price_eur_per_case']}/кор.)" for it, _ in res] + ["Весь заказ, ₽"])
    for t, k in ROWS:
        ws.append([t] + [round(r[k], 2) for _, r in res] + [round(tot(k))])
        if k in ("net", "one", "two"):
            for c in ws[ws.max_row]: c.font = Font(bold=True)
    ws.append([]); ws.append(["Поставщик", cfg.get("supplier", "")]); ws.append(["Единиц в заказе", units]); ws.append(["Маршрут", "растаможка в РБ" if by else "транзит, растаможка в РФ"])
    s2 = wb.create_sheet("Допущения"); s2.append(["Параметр", "Значение"])
    for k, v in {**cfg["fx"], **cfg["rates"], **cfg["freight"], **cfg["fixed"]}.items(): s2.append([k, v])
    for c in ws[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")
    ws.column_dimensions["A"].width = 46
    wb.save(out + ".xlsx")
except ImportError:
    print("openpyxl не установлен — Excel пропущен")

e = html.escape; f2 = lambda v: f"{v:,.2f}".replace(",", " ")
h = ["<!DOCTYPE html><html><head><meta charset='utf-8'><style>body{font-family:'DejaVu Sans',Arial;font-size:10px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #bbb;padding:2px 4px}"
     "th{background:#305496;color:#fff}td.n{text-align:right}.b{font-weight:bold;background:#f3f6fb}</style></head><body>",
     f"<h2>{e(cfg.get('title', 'Себестоимость в Москве'))}</h2><p>{e(cfg.get('supplier', ''))}. Единиц: {units:,}. Маршрут: {'растаможка в РБ' if by else 'транзит, растаможка в РФ'}. €1 = {EUR} ₽. {e(cfg.get('note', ''))}</p>".replace(",", " "),
     "<table><tr><th>₽ за единицу</th>" + "".join(f"<th>{e(it['name'])}</th>" for it, _ in res) + "<th>Заказ, ₽</th></tr>"]
for t, k in ROWS:
    h.append(f"<tr{' class=b' if k in ('net','one','two') else ''}><td>{e(t)}</td>" + "".join(f"<td class=n>{f2(r[k])}</td>" for _, r in res) + f"<td class=n>{tot(k):,.0f}</td></tr>".replace(",", " "))
h.append("</table></body></html>")
open(out + ".html", "w").write("".join(h))
chrome = next((p for p in ("/opt/pw-browsers/chromium", shutil.which("chromium") or "", shutil.which("google-chrome") or "") if p and os.path.exists(p)), None)
if chrome:
    subprocess.run([chrome, "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer", f"--print-to-pdf={os.path.abspath(out + '.pdf')}", "file://" + os.path.abspath(out + ".html")], capture_output=True, timeout=180)
for it, r in res:
    print(f"{it['name']}: без НДС {r['net']:.2f} | один НДС {r['one']:.2f} | два НДС {r['two']:.2f} ₽/ед.")
print(f"Заказ: без НДС {tot('net'):,.0f} | один НДС {tot('one'):,.0f} | два НДС {tot('two'):,.0f} ₽".replace(",", " "))
