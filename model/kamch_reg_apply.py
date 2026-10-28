"""Камчатка: номера реестровых записей ПП 719 по позициям — из найденных записей (data/kamch/registry_found.json, зеркало реестра torgs.ru + сайты заводов)
и назначения поставщиков (data/kamch/assign.json от kamch_fill.py). Номер ставится только при совпадении изготовителя и вида продукции. → data/kamch/registry.json
Перед подачей каждый номер сверить с выпиской ГИСП."""
import json, re, runpy, io, contextlib

with contextlib.redirect_stdout(io.StringIO()):
    rows = {d["tt"]: d for d in runpy.run_path("kamch_status.py")["rows"]}
AS = json.load(open("data/kamch/assign.json"))
U = "https://torgs.ru/industrial-products/"
R = {
    "ld_cel": ("10655491", "30.06.2028", "ООО «ЧелябинскСпецГражданСтрой» (ИНН 7451211335)", U + "387326/", "Краны шаровые «LD» DN10-1200 — сверить исполнение (цельносварной)"),
    "ld_nerzh": ("10657717", "30.06.2028", "ООО «ЧелябинскСпецГражданСтрой» (ИНН 7451211335)", U + "389488/", "Краны шаровые DN10-1200 (09Г2С и нержавеющая сталь) — сверить исполнение"),
    "ld_razb": ("10714005", "17.09.2028", "ООО «ЧелябинскСпецГражданСтрой» (ИНН 7451211335)", U + "440944/", "Краны шаровые 11с67п разборные (КШ.Р.) DN25-200"),
    "ld_lat": ("10826766", "04.05.2031", "ООО «ЛД Прайд» (ИНН 7449114775)", U + "542786/", "Краны шаровые латунные LD DN15-50 PN до 4,0 — изготовитель ЛД Прайд, не ЧСГС"),
    "tula": ("10612591", "12.02.2028", "АО «ПО «Тулаэлектропривод» (ИНН 7130025182)", U + "345520/", "Электроприводы многооборотные — сверить ТУ Н-В-02 с паспортом"),
    "gr": ("10837159", "13.07.2029", "ООО «Торговый Дом АДЛ» (Гранвэл)", U + "1038491/", "Затворы Гранвэл DN25-500 PN до 1,6 — тип в записи не назван"),
    "gr65": ("10357198", "15.03.2031", "ООО «Торговый Дом АДЛ» (Гранвэл)", U + "100314/", "ЗПВС DN65"),
    "gr80": ("10357205", "15.03.2031", "ООО «Торговый Дом АДЛ» (Гранвэл)", U + "100314/", "ЗПВС DN80"),
    "gr100": ("10357202", "15.03.2031", "ООО «Торговый Дом АДЛ» (Гранвэл)", U + "100314/", "ЗПВС DN100"),
    "temper": ("10717439", "23.09.2030", "ООО «Темпер»", U + "444210/", "Краны шаровые стальные TEMPER DN10-600 PN16-40"),
    "promarm": ("10655456", "24.07.2028", "ООО «ПромАрм»", U + "387291/", "Затворы поворотные дисковые ПА"),
    "ptz_zap": ("10698937", "17.02.2031", "ООО «ПТЗ» (Пермь)", U + "502692/", "Клапан запорный DN10-600 PN1,6-25"),
    "ptz_obr": ("10746921", "23.11.2028", "ООО «ПТЗ» (Пермь)", U + "502692/", "Клапан обратный подъёмный"),
}
NOTE = {"armaprom": "Арма-Пром: записи 10232737/10232729/10232727 истекли 28.08.2026 (по зеркалу) — запросить новую запись у завода",
        "lmz": "если изготовитель ЛМЗ — 10636782 (30ч6бр, до 15.05.2028); у БК-Арматуры уточнить завод",
        "granlok": "Гранлок: в реестре только тип ЗОП (Ду150 — 10618660) — уточнить у АДЛ, что CV16 = ЗОП",
        "none": "номер в открытых источниках не найден — запросить у завода"}
out = {}
for tt, d in rows.items():
    a = AS.get(tt, {}); comp = a.get("comp", ""); nm = d["name"]; txt = (nm + " " + (d["offer"] if d["act"] != "Заказать у других" else "")).lower()
    dn = (lambda m: int(m.group(1)) if m else None)(re.search(r"(?:ду|dn)\s?0*(\d+)", txt))
    k = None; note = ""
    if "ЧелябинскСпецГражданСтрой" in comp:
        if "кшх" in txt: note = "КШХ (химический) — у LD записи на такой кран не найдено; запросить у завода"
        elif "temper" in txt: note = "оригинал TEMPER — у ТЕМПЕР номер 10717439; у LD (аналог) запросить"
        else: k = "ld_nerzh" if re.search(r"нж45|нерж", txt) else "ld_lat" if re.search(r"11б27|латун", txt) else ("ld_razb" if re.search(r"разборн|кшр", txt) and dn and 25 <= dn <= 200 else "ld_cel")
    elif "Тулаэлектропривод" in comp: k = "tula"
    elif "АДЛ" in comp:
        if "гранлок" in txt: note = NOTE["granlok"]
        elif "зпвс" in txt and dn in (65, 80, 100): k = f"gr{dn}"
        elif "гранвэл" in txt or "грэнвэл" in txt: k = "gr"
    elif "ТЕМПЕР" in comp: k = "temper"
    elif "ПромАрм" in comp: k = "promarm"
    elif "Пермский" in comp or "ПТЗ" in comp: k = "ptz_obr" if "обратн" in txt else "ptz_zap"
    elif "Арма-Пром" in comp: note = NOTE["armaprom"]
    elif "БК-Арматура" in comp and "30ч6бр" in txt: note = NOTE["lmz"]
    else: note = NOTE["none"]
    if k:
        n, v, m, u, c = R[k]; out[tt] = dict(num=n, valid_to=v, man=m, src=u, note=c)
    else: out[tt] = dict(num="", valid_to="", man="", src="", note=note)
# тот же товар у текущего поставщика, изготовитель с записью (data/kamch/dnru_assign.json): Dendor → АО «Энергия», УРРД → Энерготехномаш, Гранлок → АДЛ, ЗПД → Транс-Энерго
DA = {a["tt"]: a for a in json.load(open("data/kamch/dnru_assign.json"))}
BRAND = [(r"dendor", r"Энергия"), (r"УРРД", r"Энерготехномаш"), (r"Гранлок|CV16", r"АДЛ"), (r"ЗПД|ЗДП", r"Транс-Энерго")]
for tt, a in DA.items():
    d = rows[tt]
    if a.get("main_reg") and not out[tt]["num"] and d["act"] != "Заказать у других" and any(re.search(b, d["offer"], re.I) and re.search(m, a["main"]) for b, m in BRAND):
        out[tt] = dict(num=a["main_reg"], valid_to=a["main_reg_valid"], man=re.sub(r"\s*\(.*", "", a["main"]) + f' (ИНН {a.get("main_inn")})', src=a["main_reg_url"], note=a.get("main_reg_product", "")[:150])
json.dump(out, open("data/kamch/registry.json", "w"), ensure_ascii=False, indent=1)
import collections
print("с номером:", sum(1 for x in out.values() if x["num"]), collections.Counter(x["num"] for x in out.values() if x["num"]))
