"""Справочник цен 2026 (проверенные ссылки 14.10.2026): база себестоимости по группам товаров и соли."""
import json, statistics as st
D="data/p2026/"
rows=json.load(open(D+"prices2026_clothes.json"))+json.load(open(D+"prices2026_ppe.json"))
SALTROWS=json.load(open(D+"prices2026_salt.json"))
def bad(x):
    p=(x.get("product") or "").lower()
    return ("распродаж" in p or "распродаж" in (x.get("conditions") or "").lower()
            or (x["group"]=="Комбинезон летний" and "полукомбинезон" in p))
W={"Костюм летний рабочий":1.2,"Костюм летний спасателя/МЧС":1.3,"Костюм зимний рабочий (I–III пояс)":2.6,"Костюм зимний северный (IV/особый пояс)":3.2,
   "Комбинезон летний":1.1,"Комбинезон зимний":2.4,"Куртка утепленная":1.6,"Полукомбинезон / брюки зимние":1.0,"Футболка":0.2,"Халат рабочий":0.5,
   "Жилет утепленный":0.7,"Жилет сигнальный":0.15,"Костюм сигнальный летний":1.3,"Костюм сигнальный зимний":2.6,"Костюм от нефтепродуктов летний":2.0,
   "Костюм от нефтепродуктов зимний":2.6,"Костюм КЩС (кислоты/щелочи)":2.4,"Костюм противоэнцефалитный":1.0,"Плащ влагозащитный":0.8,
   "Ботинки кожаные с МП (летние)":1.6,"Ботинки утепленные":1.9,"Сапоги ПВХ":1.4,"Сапоги утепленные / валенки":1.8,"Полусапоги":1.7,"Костюм сварщика":3.0,
   "Костюм от электродуги летний":2.0,"Костюм от электродуги зимний":3.0,"Термобельё (комплект)":0.4,"Бельё термостойкое":0.5,"Перчатки х/б с ПВХ (пара)":0.08,
   "Перчатки трикотажные":0.06,"Перчатки термостойкие/краги":0.3,"Рукавицы (пара)":0.15,"Каска":0.4,"Подшлемник":0.1,"Щиток защитный лицевой":0.4,"Перчатки резиновые":0.1}
CLIENT={"Костюм летний рабочий":3500,"Костюм летний спасателя/МЧС":3500,"Костюм зимний рабочий (I–III пояс)":5000}
G={}
for x in rows:
    if x.get("price") is None: continue
    G.setdefault(x["group"],[]).append(x)
def market(g):
    r=[x for x in G.get(g,[]) if not bad(x)]
    opt=[x["price"] for x in r if "розниц" not in (x.get("price_type") or "") or "опт" in (x.get("price_type") or "")]
    use=opt or [x["price"] for x in r]
    return st.median(use) if use else None
def base(g):
    """себестоимость для расчёта, ₽/ед. (без доставки)"""
    return CLIENT.get(g) or market(g)
def weight(g): return W.get(g,1.0)
def sources(g,n=3):
    r=[x for x in G.get(g,[]) if not bad(x)]
    r.sort(key=lambda x:(not x["date_evidence"].startswith("A"),x["price"]))
    return r[:n]
# Соль: себестоимость с доставкой до заказчика, ₽/т (по проверенным ценам 2026)
SALT2026={
 "32616451106":(5400,"МО: с доставкой 4 600–5 950 (perevalki.ru, Белкалий МКР) / 4 900–6 000 (mariva.ru); середина ≈5 400",["https://perevalki.ru/catalog/sol-tehnicheskaya/","https://mariva.ru/product-category/antigololyodnye-reagenty/sol-tehnicheskaya/","https://nsmgr.ru/reagenty/tekhnicheskaya-sol/navalom"]),
 "0851300006226000106":(8200,"Новосибирск: 7 900 ₽/т от 69 т МКР (Полевской Мрамор, отгрузка с Урала) + местная доставка ~300",["https://polevskoi-mramor.com/novosibirsk/sol-tekhnicheskaya-galit/dlya-dorog/","https://smola38.ru/promyshlennaya-khimiya/sol-tekhnicheskaya-mineralnyy-galit/"]),
 "0128600006826000407":(5800,"Струнино (~100 км от МКАД): МО с доставкой ≈5 400 + удлинение плеча ~400; Владимир/Калуга 5 000–8 200",["https://perevalki.ru/catalog/sol-tehnicheskaya/","https://tehsalt.ru/mineralnyj-galit.html","https://icemelt33.ru/texnicheskaya-sol/"]),
 "32616447519":(5600,"Волоколамск (~120 км): МО с доставкой ≈5 400 + плечо ~200",["https://perevalki.ru/catalog/sol-tehnicheskaya/"]),
 "0318300065126000256":(7000,"Краснодарский край: открытых цен 2026 нет — Москва/Центр 6 300–8 000 + доставка [оценка]",["https://solnerud.ru/prajs-list/"]),
 "0131300000626001316":(6500,"Воронеж: Центр 5 000–8 000 + доставка [оценка]",["https://tehsalt.ru/mineralnyj-galit.html","https://solnerud.ru/prajs-list/"]),
 "0142200001326018965":(6500,"Самара: Урал/ПУХК от 6 000 + доставка [оценка]",["https://puhk.info/katalog/sol/kontsentrat-mineralnyy-galit/"]),
 "0372200193426000027":(12000,"СПб, мешки 20–25 кг: МКР/мешки 10 400–15 050 (masterkrowli.ru)",["https://www.masterkrowli.ru/sol-tehnicheskaya"]),
}
