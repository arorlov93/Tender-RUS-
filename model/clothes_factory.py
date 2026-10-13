"""Пересчёт маржи по одежде по оптовым/заводским ценам РФ (13 сайтов, 14.10.2026)."""
import json, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
def P(nm,c,d):
    p=nm*(1-d); pre=p-c-p*0.008-c*0.2*60/365; return pre-max(pre,0)*0.15
def be(nm,c):
    lo,hi=-0.5,1.0
    for _ in range(60):
        m=(lo+hi)/2
        if P(nm,c,m)>0: lo=m
        else: hi=m
    return lo
# Справочник: (группа, сайт, производитель/город, модель, цена, тип цены, url)
REF=[
("Летний рабочий костюм","ivuniforma.ru","Иваново, фабрика","Р-02-3-ГР / Р-02-21-ГР / Р-02-184-ГР",(938,1715),"опт","https://www.ivuniforma.ru/catalog/rabochaya-odezhda/letnyaya-spetsodezhda/kostumy/"),
("Летний рабочий костюм","spets.ru","ГК Спецобъединение (Москва)","ОПЕРАТОР / СТРОЙМАСТЕР / ДЮРАБЛ",(3367,4984),"прайс с НДС","https://www.spets.ru/products/spetsodezhda/letnyaya/"),
("Летний рабочий костюм","trakt.ru","Тракт (Москва)","РИТМ / ШКВАЛ / МАСТАК",(3086,3746),"прайс для юрлиц с НДС","https://www.trakt.ru/product/11954/"),
("Летний костюм МЧС/спасателя","voen36.ru","Воен36 (Воронеж, собств. пр-во)","габардин / рип-стоп / ПЭВ",(3100,5995),"розница","https://voen36.ru/catalog/forma_mchs/"),
("Летний костюм МЧС/спасателя","barrakuda.biz","Барракуда (Москва)","рип-стоп / п/в",(4900,6675),"розница, есть опт","https://barrakuda.biz/formennaya-odezhda/kostyumy-mchs"),
("Летний костюм МЧС/спасателя","productcenter.ru","«Аврора» (Барнаул, фабрика)","полевой МЧС",(2490,2490),"опт «от»","https://productcenter.ru/products/45461/kostiumy-dlia-sotrudnikov-mchs"),
("Летний полевой костюм","ivsoteks.ru","Ивсотекс (Иваново)","Святогор Greta / Горка Rip-stop",(1600,2490),"опт","https://ivsoteks.ru/specodejda"),
("Летний полевой костюм","splav.ru","Сплав (СПб, фабрика)","Костюм полевой Сплав",(4600,8700),"розница","https://www.splav.ru/catalog/military-tactics/novyj-odezhda/kostyumy/letnie-kostyumy/"),
("Зимний рабочий костюм","ivuniforma.ru","Иваново, фабрика","Р-02-66 / Р-02-45 / Р-02-61",(3549,5810),"опт","https://www.ivuniforma.ru/catalog/rabochaya-odezhda/zimnjaja/kostyumy-1/"),
("Зимний рабочий костюм","specteks37.ru","СпецТекс 37 (Иваново)","№1 / №2 / №265 (IV пояс)",(2650,4600),"опт «от»","https://www.specteks37.ru/katalog/zimnyaya-spetsodezhda/"),
("Зимний рабочий костюм","spec-standart.ru","Стандарт Спецодежда (Иваново)","Сибирь / Зимний / Север (тройка)",(4158,8543),"«от»","https://spec-standart.ru/zimnyaya-specodezhda"),
("Зимний рабочий костюм","ivsoteks.ru","Ивсотекс (Иваново)","Святогор / Урал / Привал",(3520,5830),"опт","https://ivsoteks.ru/specodejda"),
("Зимний рабочий костюм","spets.ru","ГК Спецобъединение","УРЕНГОЙ / СИБИРЬ / НАДЫМ / УРАГАН",(6084,15250),"прайс с НДС","https://www.spets.ru/products/spetsodezhda/odezhda_zimnyaya/"),
("Зимний рабочий костюм","trakt.ru","Тракт","ВАХТА / ДЖОБЕР / БРИГАДИР / ФРОСТ",(3588,7793),"прайс для юрлиц с НДС","https://www.trakt.ru/product/12066/"),
("Зимний костюм северный (IV/особый пояс)","spets.ru","ГК Спецобъединение","НАДЫМ / УРАГАН-УЛЬТРА",(9760,22875),"прайс с НДС","https://www.spets.ru/products/spetsodezhda/odezhda_zimnyaya/"),
("Зимний костюм северный (IV/особый пояс)","trakt.ru","Тракт","БЕРИНГ / СЕВЕРНЫЙ ШЕЛЬФ (куртка+п/к)",(9659,23998),"прайс для юрлиц с НДС","https://www.trakt.ru/product/13086/"),
("Зимний костюм спасателя МЧС","barrakuda.biz","Барракуда","поливискоза, с капюшоном",(14461,14461),"розница","https://barrakuda.biz/formennaya-odezhda/kostyumy-mchs"),
("Зимняя куртка МЧС","voen36.ru","Воен36","уставная / удлинённая",(4800,7150),"розница","https://voen36.ru/catalog/forma_mchs/"),
("Комбинезон летний","ivuniforma.ru","Иваново, фабрика","Р-11-87 / Р-11-88",(3766,4102),"опт","https://www.ivuniforma.ru/catalog/rabochaya-odezhda/letnyaya-spetsodezhda/kombinezony/"),
("Комбинезон зимний","spets.ru","ГК Спецобъединение","ПИЛОТ зимний",(10199,10577),"прайс с НДС","https://www.spets.ru/products/spetsodezhda/odezhda_zimnyaya/"),
("Футболка","spets.ru","ГК Спецобъединение","160 г/м², кор. рукав",(488,488),"прайс с НДС","https://www.spets.ru/products/spetsodezhda/trikotazh/"),
("Футболка","trakt.ru","Тракт","Футболка синяя",(484,484),"прайс","https://www.trakt.ru/product/12012/"),
("Футболка","productcenter.ru","фабрики Иваново","мужская футболка",(211,270),"опт «от»","https://productcenter.ru/products/r-ivanovskaia-obl-168/c-ivanovo-2550/catalog-muzhskiie-futbolki-1062"),
("Халат рабочий","ivuniforma.ru","Иваново, фабрика","Р-20-104 / Р-20-251 / Р-20-256",(609,1064),"опт","https://www.ivuniforma.ru/catalog/rabochaya-odezhda/letnyaya-spetsodezhda/khalati/"),
("Халат рабочий","fabrikant-textile.ru","Фабрикант (Иваново)","халат женский смесовка",(400,400),"«от»","https://fabrikant-textile.ru/catalog/spetsodezhda/"),
("Жилет утепленный","ivuniforma.ru","Иваново, фабрика","Р-06-112 / Р-06-206",(1190,1372),"опт","https://www.ivuniforma.ru/catalog/rabochaya-odezhda/zimnjaja/"),
("Жилет утепленный","spets.ru","ГК Спецобъединение","САТУРН",(2788,2788),"прайс с НДС","https://www.spets.ru/products/spetsodezhda/odezhda_dlya_okhrannykh_struktur/"),
("Жилет сигнальный","spets.ru","ГК Спецобъединение","сигнальный / ГОСТ",(390,677),"прайс с НДС","https://www.spets.ru/products/spetsodezhda/odezhda_spetsialnaya/"),
("Жилет сигнальный","fabrikant-textile.ru","Фабрикант (Иваново)","с СОП",(260,320),"«от»","https://fabrikant-textile.ru/catalog/spetsodezhda/"),
("Костюм от нефтепродуктов","spets.ru","ГК Спецобъединение","МЕГАСТАТОЙЛ летний / зимний",(13219,24199),"прайс с НДС","https://www.spets.ru/products/spetsodezhda/odezhda_spetsialnaya/"),
("Костюм от нефтепродуктов","trakt.ru","Тракт","летний 12492 / ПЛАСТОЙЛ зимний",(6731,11253),"прайс для юрлиц с НДС","https://www.trakt.ru/product/12720/"),
("Костюм сигнальный","spets.ru","ГК Спецобъединение","РЕФЛЕКТ-1/2 летний / зимний",(6112,12780),"прайс с НДС","https://www.spets.ru/products/spetsodezhda/odezhda_spetsialnaya/"),
]
# Пересчёт лотов по «заводской» цене (нижняя граница опт. цен фабрик, если есть; иначе прайс −10%)
LOTS=[]
def lot(num,name,nm,d,items,L,note):
    C=sum(q*c for _,q,c,_ in items)
    LOTS.append(dict(num=num,name=name,nm=nm,d=d,items=items,C=C,L=L,m=(nm-C-L)/nm,pn=P(nm,C+L,0),pe=P(nm,C+L,d),be=be(nm,C+L),note=note))
lot("0358200022926000030","Ростов, обмундирование спасателей (44-ФЗ аукцион, 19.10)",6224420.7,0.17,[
 ("Костюм спасателя зимний",168,5000,"правило клиента; рынок опт 3 549–5 810 (Иваново), розн. МЧС 14 461"),
 ("Костюм спасателя летний",241,3500,"правило клиента; рынок МЧС 2 490–5 995"),
 ("Комбинезон спасателя зимний",208,7100,"Спецобъединение ПИЛОТ 10 199 −30% опт [оценка]; клиентский 5 000 — оптимистично"),
 ("Комбинезон спасателя летний",239,3766,"ivuniforma Р-11-87 опт 3 766"),
 ("Футболка форменная (с нанесением)",268,350,"Иваново опт 211–270 + нанесение ~100 [оценка]"),
],111495,"Пересчёт с реальными ценами комбинезонов (раньше приравнивали к комплектам)")
lot("32616443256","Якутия ГПС, форменная одежда (223-ФЗ, 15.10)",4999001.85,0.15,[
 ("Костюм летний полевой",405,3500,"правило клиента; рынок опт 1 600–2 490 (Ивсотекс), розн. 4 600+ (Сплав)"),
 ("Зимняя форменная одежда (IV пояс)",150,8543,"Спецстандарт «Север» (тройка) 8 543; СпецТекс №265 IV пояс 4 600; Надым 9 760"),
],150000,"Зимняя — по костюму IV пояса; если «особый пояс −45» уровня Северный шельф 24 000 — убыток")
lot("32616431081","ФБУ, спецодежда 18 позиций (223-ФЗ, 19.10)",3637476.36,0.10,[
 ("Жилет сигнальный",14,320,"Фабрикант 260–320"),
 ("Жилет утепленный",64,1190,"ivuniforma опт 1 190"),
 ("Костюм противоэнцефалитный",1,975,"Фабрикант 975 / Ивсотекс 1 250"),
 ("Костюм летний от мех. воздействий (муж+жен)",24,1715,"ivuniforma опт 1 610–1 715"),
 ("Куртка зимняя",20,2688,"ivuniforma опт 2 688"),
 ("Полукомбинезон зимний",20,958,"Спецстандарт 958"),
 ("Костюм сигнальный летний",4,6112,"Спецобъединение РЕФЛЕКТ-2"),
 ("Костюм сигнальный зимний",4,11621,"Спецобъединение РЕФЛЕКТ-2 зимний"),
 ("Костюм от нефтепродуктов зимний",31,11253,"Тракт ПЛАСТОЙЛ (Спецобъединение 24 199)"),
 ("Костюм от нефтепродуктов летний",56,6731,"Тракт (Спецобъединение 13 219)"),
 ("Плащ от воды",49,971,"Тракт 971 / Фабрикант 1 365"),
 ("Футболка красная",112,300,"Иваново опт 211–270"),
 ("Халат белый",70,735,"ivuniforma 721–770"),
 ("Халат синий",10,735,"ivuniforma 735"),
],38303,"Цен за ед. в ЕИС нет — НМЦК 3,64 млн против нашей закупки; проверить ТЗ (ткани/логотипы)")
wb=openpyxl.Workbook(); ws=wb.active; ws.title="Пересчёт лотов"
B=Font(bold=True); MONEY='#,##0" ₽"'
for x in LOTS:
    ws.append([f'{x["num"]} — {x["name"]}']); ws.cell(ws.max_row,1).font=Font(bold=True,size=12)
    ws.append(["Позиция","Кол-во","Цена закупки, ₽","Сумма, ₽","Откуда цена"])
    for c in ws[ws.max_row]: c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="305496")
    for n,q,c,s in x["items"]:
        ws.append([n,q,c,q*c,s]); ws.cell(ws.max_row,3).number_format=MONEY; ws.cell(ws.max_row,4).number_format=MONEY
    for lbl,v,f in (("Закупка итого",x["C"],MONEY),("Доставка",x["L"],MONEY),("НМЦК",x["nm"],MONEY),("Маржа при НМЦК",x["m"],"0%"),
                    ("Чистая прибыль при НМЦК",x["pn"],MONEY),(f'Чистая прибыль при снижении {x["d"]:.0%}',x["pe"],MONEY),("Безубыток до снижения",x["be"],"0%")):
        ws.append([lbl,None,None,v,x["note"] if lbl=="Закупка итого" else None]); ws.cell(ws.max_row,1).font=B; ws.cell(ws.max_row,4).number_format=f; ws.cell(ws.max_row,4).font=B
    ws.append([])
for i,w in enumerate([44,9,16,16,80],1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
for row in ws.iter_rows():
    for c in row: c.alignment=Alignment(wrap_text=True,vertical="top")
ws2=wb.create_sheet("Цены заводов и оптовиков")
ws2.append(["Группа товара","Сайт","Производитель / город","Модели","Цена от, ₽","Цена до, ₽","Тип цены","Ссылка"])
for c in ws2[1]: c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="305496")
for g,s,m,mod,(a,b),t,u in REF:
    ws2.append([g,s,m,mod,a,b,t,u]); r=ws2.max_row; ws2.cell(r,5).number_format=MONEY; ws2.cell(r,6).number_format=MONEY
    ws2.cell(r,8).hyperlink=u; ws2.cell(r,8).font=Font(color="0563C1",underline="single")
for i,w in enumerate([32,18,30,36,12,12,18,60],1): ws2.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
ws2.auto_filter.ref=f"A1:H{ws2.max_row}"; ws2.freeze_panes="A2"
wb.save("../CLOTHES_FACTORY_PRICES_1410.xlsx")
for x in LOTS: print(x["num"],round(x["C"]),f'{x["m"]:.0%}',round(x["pn"]),round(x["pe"]),f'{x["be"]:.0%}')
print(len(REF),"цен,",len({r[1] for r in REF}),"сайтов")
