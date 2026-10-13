"""Ручная сверка одежды по реальным каталогам (Тракт, Спецобъединение) + исправление Якутии."""
import json, re, statistics as st
SRC="/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/deep"
def P(nm,c,d):
    p=nm*(1-d); pre=p-c-p*0.008-c*0.2*60/365; return pre-max(pre,0)*0.15
def be(nm,c):
    lo,hi=-0.5,1.0
    for _ in range(60):
        m=(lo+hi)/2
        if P(nm,c,m)>0: lo=m
        else: hi=m
    return lo
d=json.load(open("data/deep/out_clothes_nmck.json"))
for x in d:
    if x["num"]=="32616443256":
        nm=x["nmck"]; L=150_000
        for p in x["positions"]:
            if p["item"].startswith("Зимняя"):
                p["cost_unit"]=9659.0; p["cost_src"]="https://www.trakt.ru/product/13086/ (костюм утепленный БЕРИНГ, Тракт 9 659 ₽ с НДС); пессимистично для особого пояса — СЕВЕРНЫЙ ШЕЛЬФ куртка 14 295 + п/к 9 703 = 23 998 ₽ (https://www.trakt.ru/product/133830/, https://www.trakt.ru/product/133856/)"
                p["nmck_unit"]=None
            else:
                p["nmck_unit"]=None
        C=405*3500+150*9659
        x["cost_total"]=C; x["logistics"]=L; x["margin_pct"]=(nm-C-L)/nm
        x["net_at_nmck"]=round(P(nm,C+L,0)); x["exp_disc"]=0.15; x["net_expected"]=round(P(nm,C+L,0.15)); x["breakeven_disc"]=be(nm,C+L)
        Cp=405*3500+150*23998
        x["risks"]=("ИСПРАВЛЕНО: зимние комплекты ранее не были посчитаны. База: летний 3 500 (правило клиента), зимний 9 659 (Тракт «Беринг»). "
                    f"Если нужен костюм для особого пояса (−40…−45 °C) уровня «Северный шельф» 23 998 ₽ — себестоимость {Cp+L:,.0f} ₽, маржа {(nm-Cp-L)/nm:.0%}, безубыток при снижении {be(nm,Cp+L):.0%}. "
                    "Цены за ед. в 223-ФЗ не раскрыты — уточнить ТЗ (класс защиты/пояс). Доставка в Якутию 150 тыс. [оценка].").replace(","," ")
        x["confidence"]="med"; x["decision"]="Участвовать" if be(nm,Cp+L)>0.05 else "Осторожно"
json.dump(d,open("data/deep/out_clothes_nmck.json","w"),ensure_ascii=False,indent=1)
# справочник рыночных цен одежды по каталогам
t=json.load(open(f"{SRC}/web/trakt_all.json")); seen=set(); t=[z for z in t if z.get("price") and not (z["url"] in seen or seen.add(z["url"]))]
GROUPS=[("Костюм летний рабочий (куртка+брюки/п-к)",r"^костюм",r"утепл|зимн|дуг|сварщ|нефт|кислот|сигнал|термич|электро|ОПЗ|влагозащ|химич|пчел|энцеф|лесоруб|медиц|пищев|повышенн"),
        ("Костюм зимний утепленный (I–III пояс)",r"^костюм.*(зимн|утепл)",r"дуг|сварщ|нефт|кислот|сигнал|ОПЗ|электро|термич|СЕВЕР"),
        ("Северная серия (особый/IV пояс)",r"СЕВЕРН",r"наушник"),
        ("Костюм от электродуги",r"дуг",None),
        ("Костюм сварщика",r"сварщ",None),
        ("Костюм сигнальный",r"^костюм.*сигнал",None),
        ("Куртка утепленная",r"^куртка.*утепл",r"дуг|сигнал"),
        ("Полукомбинезон / брюки утепленные",r"^(полукомбинезон|брюки).*утепл",r"дуг"),
        ("Ботинки утепленные",r"^ботинки.*(утепл|мех)",None),
        ("Ботинки летние",r"^ботинки",r"утепл|мех"),
        ("Сапоги",r"^сапоги",None),
        ("Бельё нательное / термобельё",r"бель",None),
        ("Футболка / поло",r"^(футболка|поло)",None),
        ("Жилет",r"^жилет",None),
        ("Перчатки",r"^перчатки",None),
        ("Рукавицы",r"^рукавицы",None),
        ("Каска / подшлемник",r"^(каска|подшлемник)",None),
        ("Головной убор",r"^(шапка|кепи|бейсболк|берет)",None)]
bench=[]
for g,pat,neg in GROUPS:
    r=[z for z in t if re.search(pat,z["name"],re.I) and not (neg and re.search(neg,z["name"],re.I))]
    if not r: continue
    r.sort(key=lambda z:z["price"]); ps=[z["price"] for z in r]
    ex=[r[0],r[len(r)//2],r[-1]]
    bench.append({"Группа":g,"Моделей в каталоге":len(r),"Мин., ₽":round(min(ps)),"Медиана, ₽":round(st.median(ps)),"Макс., ₽":round(max(ps)),
                  "Примеры (мин / медиана / макс)":"\n".join(f'{z["name"][:70]} — {z["price"]:,.0f} ₽ — {z["url"]}'.replace(","," ") for z in ex),
                  "Источник":"Каталог ООО «Тракт» (trakt.ru), цены для юрлиц с НДС, 13.10.2026; опт. скидка обычно 5–15% от объёма"})
bench.insert(0,{"Группа":"СЕБЕСТОИМОСТЬ КЛИЕНТА (правило)","Моделей в каталоге":None,"Мин., ₽":3500,"Медиана, ₽":None,"Макс., ₽":5000,
                "Примеры (мин / медиана / макс)":"Летний комплект 3 500 ₽, зимний 5 000 ₽ — соответствует нижней трети каталога Тракт (летний медиана ~3 500, зимний типовой 3 588–7 793)","Источник":"данные клиента"})
json.dump(bench,open("data/deep/bench_clothes_market.json","w"),ensure_ascii=False,indent=1)
for b in bench: print(b["Группа"],b["Моделей в каталоге"],b["Мин., ₽"],b["Медиана, ₽"],b["Макс., ₽"])
y=[x for x in d if x["num"]=="32616443256"][0]; print(y["margin_pct"],y["net_expected"],y["breakeven_disc"],y["decision"]); print(y["risks"])
