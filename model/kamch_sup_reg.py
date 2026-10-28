"""Камчатка 32616346943: по каждому текущему поставщику — его предложение построчно (как в счёте/КП) + графы для заполнения:
завод-изготовитель, ИНН, страна, № реестровой записи ПП 719 / ЕАЭС, срок, продукция в записи; при отсутствии записи — российская замена.
Известные номера предзаполнены (data/kamch/registry.json). → ../KAMCHATKA_REESTR_POSTAVSHIKI/*.xlsx"""
import json, os, re, shutil, runpy, io, contextlib, openpyxl
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Font, PatternFill, Border, Side
from tidy import tidy

with contextlib.redirect_stdout(io.StringIO()):
    S = runpy.run_path("kamch_status.py")
rows = S["rows"]; INV = S["INV"]
REG = json.load(open("data/kamch/registry.json"))
kp = lambda f: json.load(open(f"data/kamch/kp/{f}.json"))
by_z = {(d["grp"], d["z"]): d for d in rows}
SUP = []
# ДН.РУ — два счёта, строки как в счёте (основная позиция + КОФ)
for grp, f in (("Задвижки", "dnru_819892_v2"), ("Клапаны", "dnru_819896_v2")):
    d0 = kp(f); L = []
    for x in d0["items"]:
        r = by_z.get((grp, x["z"]), {})
        L.append(dict(line=f'{d0["invoice"].split(" (")[0]}, № {x["z"]}', tt=r.get("tt", ""), tz=r.get("name", ""), offer=x["name"], art=x.get("art", ""), q=x["qty"],
                      price=round(x["sum_vat"] / x["qty"], 2) if x.get("qty") else None))
    SUP.append(("ДН.РУ", d0["supplier"], "", L)) if grp == "Задвижки" else SUP[-1][3].extend(L)
tz = {d["tt"]: d["name"] for d in rows}; qq = {d["tt"]: d["q"] for d in rows}
for name, f, contact in (("ТД АДЛ (Гранвэл)", "adl_utd112759", "+7 (495) 937-89-68, info@adl.ru"), ("Пензапромарматура (КИПИКА)", "kipika_0110", "+7 (8412) 909-200, b18bk@11b18bk.ru"),
                         ("Лидер-Энерго", "lider-energo_2022", "+7 (8482) 21-66-32, trans-energo@bk.ru"), ("Тулаэлектропривод", "tula_10-160", "+7 (4872) 72-47-16, td@tulaprivod.ru")):
    d0 = kp(f); L = []
    for x in d0["items"]:
        t = x.get("tt") or x.get("z")
        L.append(dict(line=d0["invoice"], tt=t, tz=tz.get(t, ""), offer=x["name"], art=re.search(r"арт\. (\d+)", x["name"]).group(1) if "арт." in x["name"] else "",
                      q=x.get("qty") or qq.get(t), price=x.get("unit_vat")))
    SUP.append((name, d0["supplier"], contact, L))

FILL = ["Завод-изготовитель (как в реестре)", "ИНН изготовителя", "Страна происхождения", "№ реестровой записи ПП 719 / евразийского реестра", "Срок действия записи",
        "Наименование продукции в записи (из выписки ГИСП)", "Если записи нет: российская замена — марка, завод, № записи, цена с НДС"]
thin = Side(style="thin", color="BFBFBF"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin); Y = PatternFill("solid", fgColor="FFF2CC"); G = PatternFill("solid", fgColor="E2EFDA")
OUT = "../KAMCHATKA_REESTR_POSTAVSHIKI"; shutil.rmtree(OUT, ignore_errors=True); os.makedirs(OUT)
for k, (name, sup, contact, L) in enumerate(SUP, 1):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Реестр по позициям"
    ws.append([f"ООО «ГИДРОПАС» → {sup}: подтверждение страны происхождения по вашему предложению (ПАО «Камчатскэнерго», закупка 32616346943)"]); ws["A1"].font = Font(bold=True, size=12)
    ws.append(["Заказчик применяет ограничение по ПП 1875: в заявке по каждой позиции указываются изготовитель, страна и ДЕЙСТВУЮЩИЙ номер реестровой записи ПП 719 (для ЕАЭС — евразийского реестра)."])
    ws.append(["Изготовитель должен совпадать с записью, иначе товар считается иностранным. Просим заполнить жёлтые графы и приложить выписки из ГИСП. Зелёным — уже найдено нами в открытом реестре, просим подтвердить."])
    ws.append(["Позиции иностранного производства или без записи в реестре — просим предложить российскую замену (последняя графа)."])
    ws.append(["№", "Строка вашего счёта / КП", "Табл./№ ТЗ заказчика", "Позиция ТЗ заказчика", "Ваше предложение (как в счёте)", "Артикул", "Кол-во", "Цена за ед. с НДС (ваша), ₽", *FILL])
    for c in ws[5]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")
    for j in range(9, 9 + len(FILL)): ws.cell(5, j).fill = PatternFill("solid", fgColor="BF8F00")
    for i, x in enumerate(L, 1):
        r = REG.get(x["tt"], {}) if not re.search(r"КОФ|комплект ответных", x["offer"], re.I) else {}
        pre = [None] * len(FILL)
        if r.get("num") and r.get("man"):
            inn = re.search(r"ИНН (\d+)", r["man"]); pre[0] = re.sub(r"\s*\(ИНН \d+\)", "", r["man"]); pre[1] = inn.group(1) if inn else None
            pre[2] = "Россия"; pre[3] = r["num"]; pre[4] = r["valid_to"]; pre[5] = r.get("note")
        ws.append([i, x["line"], x["tt"], ILLEGAL_CHARACTERS_RE.sub("", x["tz"]), ILLEGAL_CHARACTERS_RE.sub("", x["offer"]), x["art"], x["q"], x["price"]] + pre)
        X = ws.max_row; ws.cell(X, 8).number_format = '#,##0.00" ₽"'
        for j in range(9, 9 + len(FILL)): ws.cell(X, j).fill = G if pre[3] and j <= 14 else Y
        for c in ws[X]: c.border = BOX
    for i, w in enumerate([4, 20, 9, 34, 44, 13, 7, 13, 30, 13, 13, 22, 12, 30, 40], 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = "E6"; ws.auto_filter.ref = f"A5:O{ws.max_row}"
    tidy(wb, row_h=52)
    fn = f"{k:02d}_" + re.sub(r"[^\wА-Яа-яЁё-]+", "_", name)[:30].strip("_") + "_реестр.xlsx"
    wb.save(f"{OUT}/{fn}")
    print(fn, len(L), "строк, предзаполнено", sum(1 for x in L if REG.get(x["tt"], {}).get("num")), "|", contact or sup)
