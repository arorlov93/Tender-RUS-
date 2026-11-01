"""Камчатка: вписать в ОФИЦИАЛЬНУЮ форму КП заказчика (data/kamch/forms/kp_form_official.xlsx) зелёные столбцы 6–8:
страна, наименование реестра и № реестровой записи ПП 719, производитель (data/kamch/forms/kp_form_map.json). Правка XML — форма заказчика
(проверки данных, форматы, формулы) не меняется. Позиции без записи остаются «…». → ../KAMCHATKA_ZAYAVKA/02a_КП_официальная_форма_заполнено.xlsx"""
import json, zipfile
from lxml import etree
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
FM = json.load(open("data/kamch/forms/kp_form_map.json"))
z = zipfile.ZipFile("data/kamch/forms/kp_form_official.xlsx"); root = etree.fromstring(z.read("xl/worksheets/sheet1.xml"))
rows = {int(r.get("r")): r for r in root.find(f"{{{NS}}}sheetData").findall(f"{{{NS}}}row")}
cn = lambda ref: sum((ord(ch) - 64) * 26 ** i for i, ch in enumerate(reversed("".join(c for c in ref if c.isalpha()))))
n = 0
for tt, v in FM.items():
    if not v["val"]: continue
    row = rows[v["row"]]
    for col, text in zip(("H", "I", "J"), v["val"]):
        ref = f'{col}{v["row"]}'; c = row.find(f'{{{NS}}}c[@r="{ref}"]')
        if c is None:
            c = etree.SubElement(row, f"{{{NS}}}c"); c.set("r", ref)
            cells = sorted(row.findall(f"{{{NS}}}c"), key=lambda e: cn(e.get("r")))
            for e in cells: row.remove(e)
            for e in cells: row.append(e)
        for ch in list(c): c.remove(ch)
        c.set("t", "inlineStr"); t = etree.SubElement(etree.SubElement(c, f"{{{NS}}}is"), f"{{{NS}}}t"); t.text = text
    n += 1
out = "../KAMCHATKA_ZAYAVKA/02a_КП_официальная_форма_заполнено.xlsx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as w:
    for it in z.infolist():
        w.writestr(it, etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True) if it.filename == "xl/worksheets/sheet1.xml" else z.read(it.filename))
print("заполнено позиций", n, "без записи", sum(1 for v in FM.values() if not v["val"]))
