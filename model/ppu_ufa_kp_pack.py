"""ППУ Уфа: файл для рассылки — лист «Заводы» (только производители, контакты сверены с сайтами) + листы спецификации из PPU_UFA_SPEC_KP.xlsx."""
import json, re, openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

C = json.load(open("data/ppu_ufa/ppu_center.json")) + json.load(open("data/ppu_ufa/ppu_ural.json"))
thin = Side(style="thin", color="BFBFBF"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin); M = '#,##0" ₽"'
QTY = {"108": 3, "377": 2, "426": 3, "530": 1}
CHK = "сайт, 23.10.2026: главная (шапка/подвал) и стр. контактов"
# заводы из первого расчёта (data/ppu_prices_1910.json) — контакты перепроверены 23.10 на сайтах
EXTRA = [dict(company="ООО ПКФ «ТЕПЛО»", plant_city="Ярославль", distance_to_ufa_km=1300, phone="+7 (4852) 91-96-10; +7 (4852) 91-96-22", email="info@pkfteplo.ru", site="ppu76.ru",
              contact_check="сайт ppu76.ru, 23.10.2026: шапка и подвал главной", prices={"108": 6907.4, "377": 53598.4, "426": 55110.3, "530": 91102.7}),
         dict(company="ПрофТехноПласт (завод ПНД и ППУ)", plant_city="уточнить", distance_to_ufa_km=None, phone="8 (800) 600-52-04", email="zakaz@proftehnoplast.ru", site="proftehnoplast.ru",
              contact_check="сайт, 23.10.2026: главная страница (страница контактов не открылась)", prices={"108": 7587, "377": 55458, "426": 60393, "530": 99441}),
         dict(company="ООО ПК «Сайверхот»", plant_city="Екатеринбург", distance_to_ufa_km=525, phone="+7 (343) 383-26-98", email="sale@saverhot.ru", site="saverhot.ru",
              contact_check="сайт, 23.10.2026: главная и стр. /contacts/ совпадают", prices={"108": 10222, "377": 63274, "426": 72561, "530": 107789})]
F = {}
for x in C:
    if not str(x["is_manufacturer"]).startswith("да") or "ТЭМПО" in x["company"]: continue  # ИТЗ — только оцинк./металлополимерная оболочка
    if x["item"] == "108" and ("Нижегородский" in x["company"] or "«НИК»" in x["company"]): continue  # на сайте 108×4,0/5,0, в ТЗ 4,5
    d = F.setdefault(x["company"], dict(company=x["company"], plant_city=x.get("plant_city"), distance_to_ufa_km=x.get("distance_to_ufa_km"), phone=x.get("phone"), email=x.get("email"),
                                        site=x.get("site"), contact_check=x.get("contact_check"), prices={}))
    if isinstance(x["price"], (int, float)) and x["item"] in QTY: d["prices"][x["item"]] = x["price"]
F = list(F.values()) + EXTRA
for d in F:
    d["kit"] = sum(d["prices"][k] * q for k, q in QTY.items()) if all(k in d["prices"] for k in QTY) else None
F.sort(key=lambda d: (-len(d["prices"]), d["kit"] or sum(d["prices"].values()), d["distance_to_ufa_km"] or 9999))

wb = openpyxl.load_workbook("../PPU_UFA_SPEC_KP.xlsx")
ws = wb.create_sheet("Заводы", 0); wb.active = 0
ws.append(["Заводы-производители отводов ППУ-ПЭ — для запроса КП (поставка в г. Уфа)"]); ws["A1"].font = Font(bold=True, size=13)
ws.append(["Стоимость комплекта (9 шт: 108×3, 377×2, 426×3, 530×1) — по прайсам на сайтах заводов, октябрь 2026, с НДС, без доставки. Контакты — со страниц официальных сайтов, сверены в двух местах сайта."])
H = ["№", "Завод", "Город", "До Уфы, км", "Стоимость комплекта, ₽", "Телефон", "E-mail", "Сайт", "Где проверен контакт",
     "Отправлен запрос (дата)", "Ответ / цена КП"]
ws.append(H)
for c in ws[3]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
for i, d in enumerate(F, 1):
    p = d["prices"]
    km = re.search(r"\d[\d\s]*", str(d["distance_to_ufa_km"] or "")); city = re.split(r"[(;,]", str(d["plant_city"] or ""))[0].strip()[:28]
    ws.append([i, d["company"], city, int(km.group(0).replace(" ", "")) if km else None, d["kit"] or "по запросу", d["phone"], d["email"], d["site"],
               str(d["contact_check"] or "")[:140], None, None])
    ws.cell(ws.max_row, 5).number_format = M
    for c in ws[ws.max_row]: c.border = BOX; c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[ws.max_row].height = 48
ws.row_dimensions[3].height = 44
for i, w in enumerate([4, 34, 16, 9, 16, 28, 26, 18, 36, 14, 18], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.append([])
for t in ["«По запросу» — на сайте нет цен на все 4 позиции (у НЗТИ и НИК нет 108×4,5 — в ТЗ стенка 4,5 мм).",
          "НЗТИ: на карточках «Отвод ППУ ГОСТ 10704» — в запросе требовать крутоизогнутый отвод по ГОСТ 17375/17380 (как в ТЗ).",
          "Спецификация для запроса — следующие листы: «Спецификация», «ТЗ дословно», «Документы к отгрузке»."]:
    ws.append([t]); ws.merge_cells(start_row=ws.max_row, start_column=1, end_row=ws.max_row, end_column=11)
ws.freeze_panes = "C4"
ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0; ws.sheet_properties.pageSetUpPr.fitToPage = True
wb.save("../PPU_UFA_ZAVODY_SPEC.xlsx")
print(len(F), "заводов;", sum(1 for d in F if d["kit"]), "с ценами")
for d in F: print(d["company"][:40], d["kit"], d["phone"][:30] if d["phone"] else None)
