"""Камчатка 32616346943: построчная перепроверка всех счетов/КП против ТЗ (data/kamch/check/out_*.json). → ../KAMCHATKA_PROVERKA_SCHETOV.xlsx"""
import json, os, re, collections, openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side
from tidy import tidy
R = []
for f in ("data/kamch/check/out_dnru.json", "data/kamch/check/out_other.json"):
    if os.path.exists(f):
        for x in json.load(open(f)): x.setdefault("supplier", "ДН.РУ"); R.append(x)
IN = {x["tt"]: x for f in ("data/kamch/check/in_dnru.json", "data/kamch/check/in_other.json") for x in json.load(open(f))}
vk = lambda v: "не соответствует" if str(v).lower().startswith("не соотв") else "при условии" if "услов" in str(v).lower() else "соответствует" if str(v).lower().startswith("соотв") else str(v)[:30]
COL = {"соответствует": "C6EFCE", "при условии": "FFEB9C", "не соответствует": "F8CBAD"}
thin = Side(style="thin", color="BFBFBF"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin); M = '#,##0" ₽"'
head = lambda ws: [setattr(c, "font", Font(bold=True, color="FFFFFF")) or setattr(c, "fill", PatternFill("solid", fgColor="305496")) for c in ws[ws.max_row] if c.value is not None]
txt = lambda v: "; ".join(map(str, v)) if isinstance(v, list) else str(v or "")
wb = openpyxl.Workbook(); sv = wb.active; sv.title = "Сводка"
sv.append(["Камчатскэнерго 32616346943 — построчная перепроверка счетов против ТЗ"]); sv["A1"].font = Font(bold=True, size=12)
sv.append(["Поставщик", "Позиций", "Соответствует", "При условии", "Не соответствует", "Сумма КП, ₽", "Сумма «не соответствует», ₽"]); head(sv)
by = collections.defaultdict(list)
for x in R: by[x["supplier"]].append(x)
for s, L in sorted(by.items(), key=lambda z: -len(z[1])):
    c = collections.Counter(vk(x["verdict"]) for x in L)
    sm = sum((IN.get(x["tt"], {}).get("price_vat") or 0) * float(IN.get(x["tt"], {}).get("tz_qty") or 0) for x in L)
    bad = sum((IN.get(x["tt"], {}).get("price_vat") or 0) * float(IN.get(x["tt"], {}).get("tz_qty") or 0) for x in L if vk(x["verdict"]) == "не соответствует")
    sv.append([s, len(L), c["соответствует"], c["при условии"], c["не соответствует"], sm, bad]); sv.cell(sv.max_row, 6).number_format = M; sv.cell(sv.max_row, 7).number_format = M
c = collections.Counter(vk(x["verdict"]) for x in R)
sv.append(["ИТОГО", len(R), c["соответствует"], c["при условии"], c["не соответствует"]]); [setattr(z, "font", Font(bold=True)) for z in sv[sv.max_row]]
for row in sv.iter_rows(min_row=2):
    for z in row:
        if z.value is not None: z.border = BOX
for i, w in enumerate([26, 9, 14, 12, 16, 16, 20], 1): sv.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
for s, L in sorted(by.items(), key=lambda z: -len(z[1])):
    ws = wb.create_sheet(re.sub(r"[^\wА-Яа-я ]", "", s)[:28])
    ws.append(["Табл./№", "Позиция ТЗ", "Что в счёте", "Кол-во ТЗ", "Кол-во в счёте", "Цена КП, ₽/шт", "Вердикт", "Несоответствия / что не указано", "Что запросить у поставщика"]); head(ws)
    for x in sorted(L, key=lambda x: (["не соответствует", "при условии", "соответствует"].index(vk(x["verdict"])) if vk(x["verdict"]) in COL else 3, x["tt"])):
        bad = "; ".join(f'{c.get("field")}: ТЗ «{c.get("tz")}» — счёт «{c.get("invoice")}» ({c.get("ok")})' for c in (x.get("checks") or []) if str(c.get("ok")) != "да")
        ws.append([x["tt"], x.get("tz_name"), x.get("offer_line"), x.get("qty_tz"), x.get("qty_inv"), IN.get(x["tt"], {}).get("price_vat"), txt(x.get("verdict"))[:200], bad[:900], txt(x.get("what_to_request"))[:400]])
        X = ws.max_row; ws.cell(X, 6).number_format = M; ws.cell(X, 7).fill = PatternFill("solid", fgColor=COL.get(vk(x["verdict"]), "FFFFFF"))
        for z in ws[X]: z.border = BOX
    for i, w in enumerate([8, 36, 40, 8, 9, 12, 26, 70, 46], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = "C2"; ws.auto_filter.ref = f"A1:I{ws.max_row}"
tidy(wb, row_h=60); wb.save("../KAMCHATKA_PROVERKA_SCHETOV.xlsx"); print(c, {s: len(L) for s, L in by.items()})
