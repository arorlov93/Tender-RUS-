"""Маржа по ТЗ партии 02.11: НМЦ заказчика против рыночных закупочных цен (с НДС; мы на УСН).
Позиции — data/zakupki_0211/tz/items/*.json; цены рынка — data/zakupki_0211/tz/pricing/g*_out.json (агенты, ноябрь 2026);
МАЗ — минимальная цена поставщика из обоснования НМЦ заказчика. Где НМЦ ед. неизвестна (только общая НМЦ) — маржа = (НМЦ − Σ рынок × кол-во) / НМЦ.
Где цены собраны по части позиций (топ по стоимости) — маржа по покрытой части. → ../TZ_0211_MARGIN.xlsx"""
import json, glob, os, openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from tidy import tidy

D = "data/zakupki_0211/tz"; META = json.load(open(f"{D}/items/_meta.json"))
PR = {}
for f in glob.glob(f"{D}/pricing/g*_out.json"):
    for code, L in json.load(open(f)).items():
        for x in L: PR[(code, int(x["n"]))] = x
thin = Side(style="thin", color="BFBFBF"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin); M = "#,##0"
wb = openpyxl.Workbook(); sv = wb.active; sv.title = "Итог"
sv.append(["ТЗ партии 02.11 — маржа по рыночным ценам (закупка с НДС; без доставки и торга, затем при −10 %)"]); sv["A1"].font = Font(bold=True, size=12)
sv.append(["Закупка", "№", "Подача", "НМЦ, ₽", "Позиций в ТЗ", "Позиций с ценой рынка", "Доля НМЦ, проверенная ценами", "НМЦ проверенной части, ₽", "Рынок по ней, ₽", "Маржа", "Маржа при −10 %", "Вывод"])
for c in sv[2]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True)
res = []
for code, m in META.items():
    L = json.load(open(f"{D}/items/{code}.json"))
    ws = wb.create_sheet(code[:31]); ws.append(["№", "Позиция", "Ед.", "Кол-во", "НМЦ ед., ₽", "Рынок ед., ₽", "Маржа ед.", "Совпадение", "Источник", "Ссылка"])
    for c in ws[1]: c.font = Font(bold=True)
    nm = mk = 0.0; nf = 0
    for x in L:
        p = PR.get((code, x["n"]), {}); mp = x.get("src_min") if code == "MAZ_metro" else p.get("price"); q = x.get("qty") or 0; n = x.get("nmc_unit")
        ws.append([x["n"], x["name"], x.get("unit"), q, n, mp, (n - mp) / n if (n and mp) else None, "мин. цена поставщика из обоснования НМЦ" if code == "MAZ_metro" else p.get("match"), p.get("source"), p.get("url")])
        X = ws.max_row; ws.cell(X, 7).number_format = "0%"
        for j in (5, 6): ws.cell(X, j).number_format = "#,##0.00"
        if mp: nf += 1; mk += mp * q; nm += (n or 0) * q
    for i, w in enumerate([5, 55, 8, 8, 12, 12, 8, 30, 20, 40], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    total = m.get("nmc_total") or m["nmc_sum"]
    if m["nmc_sum"] == 0:  # НМЦ ед. неизвестны — сравниваем с общей НМЦ, только если цены найдены по всем позициям
        cov = nf / len(L) if L else 0; nm_c = total if nf == len(L) else None
    else:
        nm_c = nm; cov = nm / m["nmc_sum"] if m["nmc_sum"] else 0
    mg = (nm_c - mk) / nm_c if nm_c else None; mg10 = (nm_c * 0.9 - mk) / (nm_c * 0.9) if nm_c else None
    part = f"нет НМЦ ед.; рынок по {nf}/{len(L)} найденным = {mk/total:.0%} от общей НМЦ" if (mg is None and nf and total) else None
    verdict = (part or "нет данных") if mg is None else "ИДЁМ (≥20 %)" if mg10 >= 0.2 else "на грани (0–20 %)" if mg10 >= 0 else "НЕ ИДЁМ (убыток)"
    res.append([m["title"], m.get("eis"), m.get("deadline"), total, len(L), nf, cov, nm_c, mk, mg, mg10, verdict])
for r in sorted(res, key=lambda r: -(r[10] if r[10] is not None else -9)):
    sv.append(r); X = sv.max_row
    for j in (4, 8, 9): sv.cell(X, j).number_format = M
    for j in (7, 10, 11): sv.cell(X, j).number_format = "0%"
    sv.cell(X, 12).fill = PatternFill("solid", fgColor="C6EFCE" if r[11].startswith("ИДЁМ") else "FFF2CC" if r[11].startswith("на грани") else "F8CBAD" if r[11].startswith("НЕ") else "D9D9D9")
    for c in sv[X]: c.border = BOX; c.alignment = Alignment(wrap_text=True, vertical="top")
sv.append([]); sv.append(["Ранее посчитано по извещениям ЕИС (позиции и цены): соль Усинск 69 т — 53 % при −10 %; соль МАУ ЖКХ 30 т — 55 %; труба кабельная (Москва) — 30–60 %; новогоднее оформление — 30–50 %."])
for i, w in enumerate([46, 18, 14, 14, 9, 9, 10, 14, 14, 8, 9, 20], 1): sv.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
tidy(wb, row_h=28); wb.save("../TZ_0211_MARGIN.xlsx")
for r in res: print(f"{r[0][:50]:50} НМЦ {r[3]:>12,.0f} цен {r[5]}/{r[4]} покрытие {r[6]:.0%} маржа {r[9] if r[9] is None else round(r[9],3)} −10%: {r[10] if r[10] is None else round(r[10],3)} {r[11]}")
