"""ТЗ партии 02.11 (загружены пользователем, data/zakupki_0211/tz/x/*): единый список позиций по каждой закупке
→ data/zakupki_0211/tz/items/<код>.json: [{n, name, spec, unit, qty, nmc_unit (с НДС, если известна), src_min (мин. цена поставщика из обоснования НМЦ, с НДС)}]"""
import json, os, re, glob, subprocess, openpyxl
from docx import Document

X = "data/zakupki_0211/tz/x"; OUT = "data/zakupki_0211/tz/items"; os.makedirs(OUT, exist_ok=True)
num = lambda v: (lambda s: float(s) if re.fullmatch(r"-?\d+(\.\d+)?", s) else None)(str(v).replace(" ", "").replace(",", ".").rstrip("*")) if v not in (None, "", "-") else None
rows = lambda f, i=0: list(openpyxl.load_workbook(f, data_only=True, read_only=True).worksheets[i].iter_rows(values_only=True))
META = {}

def save(code, items, **meta):
    json.dump(items, open(f"{OUT}/{code}.json", "w"), ensure_ascii=False, indent=0)
    META[code] = dict(meta, n=len(items), nmc_sum=round(sum((x.get("nmc_unit") or 0) * (x.get("qty") or 0) for x in items), 2))

# 5044/КО Газпром теплоэнерго МО — электротехника (ВПП: цена без НДС, НДС 22 %)
R = rows(f"{X}/f6d84e33/ТЗ Приложение №1 - ВПП (электротехническая пр-ция 1-2027).xlsx")
save("5044_GTE_MO_electro", [dict(n=int(r[0]), name=str(r[1]).strip(), spec=str(r[14] or "")[:400], unit=r[9], qty=num(r[2]), nmc_unit=round(num(r[10]) * 1.22, 2))
                             for r in R[5:] if r[0] and str(r[0]).isdigit() and num(r[10])], eis="32616526616", title="Газпром теплоэнерго МО — электротехническая продукция (5044/КО)", deadline="09.11 10:00")
# 5036/КО Газпром теплоэнерго Киров — КИПиА
R = rows(f"{X}/5d4e47e7/5036-КО/Ведомость поставки.xlsx")
save("5036_GTE_Kirov_KIP", [dict(n=int(r[0]), name=str(r[1]).strip(), spec=str(r[8] or "")[:400], unit=r[3], qty=num(r[2]), nmc_unit=num(r[7]))
                            for r in R[5:] if r[0] and str(r[0]).isdigit() and num(r[7])], eis="32616475914", title="Газпром теплоэнерго Киров — КИПиА (5036/КО)", deadline="06.11 (перенос)", nmc_total=4270000, note="НМЦ — прайс-лист единичных расценок; лимит договора 4,27 млн")
# Т Плюс Кировский — канцтовары (ОЗП 50300-18-08/1224)
R = rows(f"{X}/073bcdda/073bcdda-1.______________1/Приложение №5 - Сведения о НМЦ.xlsx")
save("TPlus_Kirov_kanc", [dict(n=int(r[0]), name=str(r[1]).strip(), spec="", unit=r[2], qty=num(r[3]), nmc_unit=num(r[5]))
                          for r in R[5:] if r[0] and str(r[0]).isdigit() and num(r[5])], eis="32616459808", title="Т Плюс Кировский — канцелярские товары", deadline="06.11 14:00")
# 43834 Россети Урал — трубы для кабельных линий (цена без/с НДС в колонках 43/44)
R = rows(f"{X}/3c3aa7fc/3c3aa7fc-_____________________43834/Прил. к ИК, Тех.часть, проект договора/Приложение №1.1. к ТЗ - Расчет НМЦ.xlsx")
save("43834_Rosseti_Ural_pipes", [dict(n=int(r[0]), name=str(r[4]).strip(), spec=str(r[6] or "")[:400], unit=r[14], qty=num(r[16]), nmc_unit=num(r[44]))
                                  for r in R[8:] if r and r[0] and str(r[0]).isdigit() and num(r[44])], eis="32616531898", title="Россети Урал — трубы для прокладки кабельных линий (43834)", deadline="10.11 09:00")
# Смоленскрегионтеплоэнерго — электроматериалы
R = rows(f"{X}/cc403228/cc403228-_______________________.___30.10.2026_..xlsx")
save("Smolensk_electro", [dict(n=int(r[0]), name=str(r[1]).strip(), spec=str(r[2] or "")[:400], unit=r[3], qty=num(r[4]), nmc_unit=num(r[5]))
                          for r in R[9:] if r[0] and str(r[0]).isdigit() and num(r[5])], eis="32616488988", title="Смоленскрегионтеплоэнерго — электроматериалы", deadline="06.11 10:00")
# Метрополитен СПб — запчасти МАЗ: в расчёте НМЦ 3–4 цены поставщиков (с НДС)
R = rows(f"{X}/814c056f/814c056f-__________1_______________1________2______________________________________________________.xlsx")
it = []
for r in R[8:]:
    if r[1] and str(r[1]).isdigit() and num(r[17]):
        src = [num(v) for v in r[7:11] if num(v) and not str(v).endswith("*")]
        it.append(dict(n=int(r[1]), name=str(r[2]).strip(), spec=str(r[3] or ""), unit=r[5], qty=num(r[6]), nmc_unit=num(r[17]), src_min=min(src) if src else None))
save("MAZ_metro", it, eis="32616485509", title="Петербургский метрополитен — запчасти МАЗ", deadline="06.11 10:00", nmc_total=11000000, note="НМЦ — прайс-лист; лимит договора 11,0 млн")
# УПТЖ для ППД (Татнефть) — арматура, цен в ТЗ нет
R = rows(f"{X}/260b616c/260b616c-3._________________________02.11.2026.xlsx")
save("UPTZh_Tatneft_arm", [dict(n=int(r[0]), name=str(r[1]).strip(), spec="", unit=r[2], qty=num(r[3]), nmc_unit=None) for r in R[1:] if r[0] and str(r[0]).isdigit()],
     eis="32616465624", title="УПТЖ для ППД (Татнефть) — запорная арматура", deadline="10.11 15:00", nmc_total=1719646.12)

def doc_table(f, ti, cn, cq, cu=None, cs=None, skip=1):
    d = Document(f); tb = d.tables[ti]; out = []
    for r in tb.rows[skip:]:
        c = [x.text.strip() for x in r.cells]
        q = num(c[cq].split()[0]) if c[cq] else None
        if c[cn] and q: out.append(dict(n=len(out) + 1, name=c[cn], spec=(c[cs][:400] if cs is not None else ""), unit=c[cu] if cu is not None else "", qty=q, nmc_unit=None))
    return out
save("OrVD_spec_winter", doc_table(f"{X}/502d349f/502d349f-______________________.docx", 13, 1, 4, 3, 2), eis="32616524948", title="ОрВД «Аэронавигация Юга» — зимняя спецодежда", deadline="09.11 08:00", nmc_total=7568890.98)
save("Sevteplo_KIP", doc_table(f"{X}/a9fa7d6f/a9fa7d6f-___-__________________________________.docx", 1, 1, 5, 4, 3), eis="32616530999", title="Севтеплоэнерго — аварийный запас КИПиА", deadline="09.11 10:00", nmc_total=1847918.33)
save("VOIN_ognevaya", doc_table(f"{X}/6622f0a3/6622f0a3-_______IV.____________________.DOCX", 0, 1, 5, 4, 3), eis="B2810261718575", title="Центр «ВОИН» (ЕАО) — оборудование «огневая подготовка»", deadline="06.11 11:00", nmc_total=8373178.18)
save("SC300_lift_parts", doc_table(f"{X}/9cdba1a6/Том 2_Volum 2/Техническое задание  ЗЗКЭФ-2026-116-НРБ.docx", 0, 1, 5, 4, 2, skip=2), eis="3511105", title="Трест Росспецэнергомонтаж (Росатом) — запчасти к подъёмнику SC300", deadline="10.11 10:00", nmc_total=652650)
def zv():
    d = Document(f"{X}/3f26df8f/3f26df8f-_____7_-___________________.docx"); out = []
    for r in d.tables[0].rows[1:]:
        c = [x.text.strip() for x in r.cells]
        if c[0].isdigit(): out.append(dict(n=int(c[0]), name=c[2] + " " + c[3], spec=c[5][:300], unit="шт", qty=num(c[4]), nmc_unit=None))
    return out
save("Zvezda_sudovaya_arm", zv(), eis="5590355", title="ДВЗ «Звезда» — судовая арматура", deadline="12.11 10:00", nmc_total=1271700)
def ch():
    d = Document(f"{X}/9a56ce96/9a56ce96-___________________________________________________.docx"); out = []
    for r in d.tables[0].rows:
        c = [x.text.strip() for x in r.cells]
        if num(c[-1]): out.append(dict(n=len(out) + 1, name=c[1], spec="", unit=c[2], qty=num(c[-1]), nmc_unit=None))
    return out
save("Chekhov_spec", ch(), eis="32616539033", title="Чеховское ЖКХ — спецодежда, спецобувь", deadline="11.11 11:00", nmc_total=3077114.99)
R = rows(f"{X}/e0b1be70/Приложение № 2/Приложение №2 Таблица А_0020.xlsx")
save("Med_0020", [dict(n=int(r[0]), name=str(r[1]).strip(), spec=str(r[2] or "")[:300] + " | производитель: " + str(r[4] or ""), unit=r[8], qty=num(r[7]), nmc_unit=None) for r in R[1:] if r[0] and str(r[0]).isdigit()],
     eis="ГПК611061", title="Мед. частное учреждение — расходники для взятия биоматериала (0020/КО)", deadline="12.11 14:00", nmc_total=3614241.26)
R = rows(glob.glob(f"{X}/e3389ece/*/ТЗ/Приложение №1 к ТЗ.xlsx")[0])
save("TSK_Mosenergo_chug_arm", [dict(n=int(r[0]), name=str(r[1]).strip(), spec=str(r[2] or "")[:300], unit=r[3], qty=num(r[4]), nmc_unit=num(r[5])) for r in R[3:] if r[0] and str(r[0]).isdigit() and num(r[5])],
     eis="1038631", title="ТСК Мосэнерго — запорная арматура (37-РП-ИП-2027)", deadline="13.11 09:00")
def tav():
    d = Document(glob.glob(f"{X}/da4fc84b/*.docx")[0]); out = []
    for r in d.tables[0].rows[1:]:
        c = [x.text.strip() for x in r.cells]
        if c[0].strip(".").isdigit() and num(c[4]): out.append(dict(n=len(out) + 1, name=c[1], spec=c[2][:300], unit=c[3], qty=num(c[4]), nmc_unit=None))
    return out
save("Tavrida_1222", tav(), eis="B2910261728004", title="АНО «Таврида.Арт» — пожарно-спасательное и досмотровое оборудование (№1222)", deadline="13.11 12:00", nmc_total=8930645)
json.dump(META, open(f"{OUT}/_meta.json", "w"), ensure_ascii=False, indent=1)
for k, v in META.items(): print(f"{k:28} поз {v['n']:5}  НМЦ по позициям {v['nmc_sum']:>14,.0f}  {v.get('nmc_total') or ''}")
