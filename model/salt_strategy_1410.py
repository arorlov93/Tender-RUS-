"""Соль: стратегия поставки от ближайшего источника и маржа по вариантам (справочно)."""
import openpyxl, importlib.util
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
spec=importlib.util.spec_from_file_location("sp","salt_prices_1410.py"); sp=importlib.util.module_from_spec(spec)
import sys,io; _o=sys.stdout; sys.stdout=io.StringIO(); spec.loader.exec_module(sp); sys.stdout=_o
P,be,REF=sp.P,sp.be,sp.REF
RAIL=2.0      # ₽/т·км, ж/д полувагон с арендой вагона [оценка]
TRUCK=6.0     # ₽/т·км, самосвал/бортовой 20–25 т на 50–600 км [оценка]
PEREV=500     # ₽/т, выгрузка из вагона, перевалка, затаривание/погрузка МКР [оценка]
EXW=2430      # ₽/т, ж/д норма с завода: 1 990 без НДС (esolk.ru) + НДС 22%
D=0.15
# лот, №, подача, НМЦК, т, фасовка, ближайший склад РФ (цена самовывоз, км авто), ближайший завод (название, км ж/д), последняя миля от станции, км
LOTS=[
("ДГХ, Московская обл.","32616451106","16.10 10:00",6724000,400,"не указана (уточнить)",("Склад Люберцы, тип D (nsmgr.ru)",4550,70),("Беларуськалий, Солигорск",800),40),
("Новосибирск, «Дзержинка»","0851300006226000106","16.10 05:00",4400000,400,"МКР",None,("Павлодарсоль (КЗ, Таволжан) / Алтайская соледоб. (Бурсоль)",600),30),
("Волоколамск, МО","32616447519","16.10 10:00",1593800,130,"уточнить (Галит)",("Склад Люберцы, тип D (nsmgr.ru)",4550,130),("Беларуськалий, Солигорск",800),20),
("Абинск, Краснодарский край","0318300065126000256","19.10 09:00",1368000,150,"МКР ≤1 т",None,("Баскунчак (Руссоль), Астраханская обл.",900),60),
("Воронеж","0131300000626001316","16.10 09:00",3296150,461,"МКР, помол №3",("Склад Люберцы, тип D (nsmgr.ru)",4550,520),("Баскунчак (Руссоль), Астраханская обл.",800),20),
("Струнино, Владимирская обл.","0128600006826000407","19.10",1342500,150,"МКР 0,8–1 т, Галит тип С",("Склад Люберцы, тип D (nsmgr.ru)",4550,110),("Беларуськалий, Солигорск",900),30),
("Самара (нужен опыт 24,2 млн)","0142200001326018965","19.10 11:00",121000000,11000,"МКР 1 т, Галит",None,("Илецксоль (Руссоль), Соль-Илецк",400),20),
]
def variants(L):
    name,num,end,nm,t,pack,wh,plant,lm=L
    v=[]
    v.append(("1. Казахстан (ваша цена)",2500,"2 500 ₽/т с доставкой в РФ — проверить точку доставки и мин. партию"))
    r=EXW+plant[1]*RAIL+PEREV+lm*TRUCK
    v.append(("2. Ближайший завод вагоном",r,f"{plant[0]}: {EXW} + ж/д {plant[1]} км × {RAIL:g} + перевалка {PEREV} + авто {lm} км × {TRUCK:g}"))
    if wh: v.append(("3. Ближайший склад авто",wh[1]+wh[2]*TRUCK,f"{wh[0]}: {wh[1]} + авто {wh[2]} км × {TRUCK:g}"))
    else: v.append(("3. Ближайший склад авто",None,"оптового склада с открытой ценой рядом не найдено"))
    v.append(("4. Дилер с доставкой (рынок)",8000,"7 000–8 500 у дилеров + доставка; верхняя граница"))
    return v
wb=openpyxl.Workbook(); ws=wb.active; ws.title="Итог по вариантам"
MONEY='#,##0" ₽"'; thin=Side(style="thin",color="BFBFBF")
ws.append(["Соль: себестоимость с доставкой и наша прибыль по 4 вариантам поставки (снижение 15%, чистая после УСН 15%)"]); ws["A1"].font=Font(bold=True,size=13)
ws.append(["Лучший реальный вариант — ближайший источник (вариант 2 или 3). Вариант 1 — ваша цена, нужно подтвердить. Все варианты кроме 1 — расчётные по открытым ценам."])
H=["Лот","№","Подача","НМЦК, ₽","Тонн","Цена в НМЦК, ₽/т"]
VN=["1. Казахстан (ваша цена)","2. Ближайший завод вагоном","3. Ближайший склад авто","4. Дилер с доставкой (рынок)"]
for n in VN: H+=[f"{n}: ₽/т",f"{n}: маржа",f"{n}: прибыль, ₽"]
H+=["Рекомендуемая стратегия","Прибыль при рекомендуемой, ₽"]
ws.append(H)
for c in ws[3]: c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="305496"); c.alignment=Alignment(wrap_text=True,vertical="center")
tot={n:0 for n in VN}; trec=0; strat=[]
for L in LOTS:
    name,num,end,nm,t,pack,wh,plant,lm=L
    vs=variants(L); row=[name,num,end,nm,t,nm/t]
    real=[(n,c) for n,c,_ in vs[1:3] if c]
    bestn,bestc=min(real,key=lambda x:x[1])
    for n,c,_ in vs:
        if c is None: row+=[None,None,None]; continue
        C=c*t; pe=P(nm,C,D); row+=[c,(nm-C)/nm,pe]
        if "Самара" not in name: tot[n]+=pe
    pr=P(nm,bestc*t,D); row+=[f"{bestn}: {bestc:,.0f} ₽/т".replace(","," "),pr]
    if "Самара" not in name: trec+=pr
    ws.append(row); R=ws.max_row
    ws.cell(R,2).number_format="@"; ws.cell(R,4).number_format=MONEY; ws.cell(R,6).number_format=MONEY
    for i in range(4):
        b=7+3*i; ws.cell(R,b).number_format=MONEY; ws.cell(R,b+1).number_format="0%"; ws.cell(R,b+2).number_format=MONEY
        v=ws.cell(R,b+2).value
        if v is not None: ws.cell(R,b+2).fill=PatternFill("solid",fgColor="C6EFCE" if v>300000 else ("FFEB9C" if v>0 else "FFC7CE"))
    ws.cell(R,20).number_format=MONEY; ws.cell(R,20).font=Font(bold=True)
    strat.append((L,vs,bestn,bestc,pr))
ws.append(["ИТОГО без Самары","","","","",""]+sum([["","",tot[n]] for n in VN],[])+["",trec]); R=ws.max_row
for i in range(4): ws.cell(R,9+3*i).number_format=MONEY
ws.cell(R,20).number_format=MONEY
for c in ws[R]: c.font=Font(bold=True)
for i,w in enumerate([28,21,11,14,7,11]+[10,8,14]*4+[34,16],1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
for row in ws.iter_rows(min_row=4):
    for c in row: c.border=Border(top=thin,bottom=thin,left=thin,right=thin); c.alignment=Alignment(wrap_text=True,vertical="top")
ws.freeze_panes="C4"
# стратегия по лотам
w2=wb.create_sheet("Стратегия поставки")
w2.append(["Лот","Тонн / фасовка","Ближайший завод (ж/д)","Ближайший склад (авто)","Варианты: себестоимость с доставкой, ₽/т — расчёт","Рекомендация","Шаги"])
for c in w2[1]: c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="305496"); c.alignment=Alignment(wrap_text=True)
for (name,num,end,nm,t,pack,wh,plant,lm),vs,bn,bc,pr in strat:
    vtxt="\n".join(f"{n}: {('—' if c is None else f'{c:,.0f}'.replace(',',' '))} — {d}" for n,c,d in vs)
    wag=max(1,round(t/68))
    steps=(f"1) Запросить у {plant[0].split(',')[0]} цену франко-вагон + наличие МКР; 2) заказать {wag} ваг. (~68 т/ваг.) на станцию ближе к заказчику; "
           f"3) договор с перевалкой у станции (выгрузка, хранение, погрузка МКР); 4) авто {lm} км до заказчика по графику; "
           "5) параллельно запросить КП у ближайшего склада как резерв." if bn.startswith("2") else
           f"1) Запросить КП у склада ({wh[0] if wh else ''}) с доставкой до заказчика; 2) зафиксировать цену на весь объём; 3) поставки по графику заказчика самосвалом/бортом; 4) резерв — вагон от завода.")
    w2.append([name,f"{t:g} т; {pack}",f"{plant[0]} (~{plant[1]} км)",f"{wh[0]} (~{wh[2]} км)" if wh else "—",vtxt,f"{bn}: {bc:,.0f} ₽/т → прибыль {pr:,.0f} ₽".replace(","," "),steps])
for i,w in enumerate([26,18,32,30,70,30,70],1): w2.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
for row in w2.iter_rows():
    for c in row: c.alignment=Alignment(wrap_text=True,vertical="top")
# справочно
w3=wb.create_sheet("Справочно — как считали")
notes=[
("Модель прибыли","Цена = НМЦК × (1 − 15%); прибыль = цена − себестоимость − 0,8% прочих − стоимость денег 20% годовых × 60 дн.; минус УСН 15% с прибыли."),
("Цена завода вагоном",f"1 990 ₽/т без НДС, ж/д норма ~70 т (esolk.ru: Павлодарсоль, Аралтуз, Уралкалий, Белкалий, Тыреть) → {EXW} ₽/т с НДС 22% (для УСН НДС — расход)."),
("Ж/д тариф",f"{RAIL:g} ₽/т·км с арендой полувагона [оценка; точный тариф — расчёт РЖД/оператора по станциям]. Ставки на полувагоны в 2026 выросли >40%, вагонов не хватает — закладывать запас по сроку 10–20 дней."),
("Перевалка",f"{PEREV} ₽/т: выгрузка из вагона, временное хранение, затаривание в МКР/погрузка [оценка]."),
("Авто",f"{TRUCK:g} ₽/т·км для 50–600 км [оценка]; по МО — 8–11 тыс. ₽ за машину 20–25 м³ до 60 км (perevalka-tsk.ru)."),
("Склад РФ","nsmgr.ru (Люберцы): тип D 4 350–4 550 ₽/т самовывоз; Уралкалий тип A 6 500; Илецкая тип C 5 750."),
("Дилеры (рынок)","Аралтуз 7 000–7 800, Баскунчак/Белкалий 7 500–8 700, Уралкалий 8 000–8 500, Илецк/Тыреть 10 000 ₽/т без доставки (sknerud, vektor-e, perevalka-tsk)."),
("Ваша цена 2 500","Реалистична только как цена вагоном до станции в РФ (завод КЗ ~2 000 без НДС + ж/д). Если 2 500 — до станции, добавьте перевалку и авто (+700…1 500 ₽/т)."),
("Галит / ТУ","Где в ТЗ «Концентрат минеральный „Галит“» (Уралкалий, ТУ) — проверить допуск эквивалента до закупки."),
("Фасовка МКР","Если в ТЗ МКР, а вагон идёт навалом — затаривание на перевалке (входит в 500 ₽/т) или брать у завода сразу в МКР (+300–600 ₽/т) [оценка]."),
("Срок","Подача 16–19.10; поставки обычно ноябрь–март по заявкам → держать запас на перевалке у станции ближе к заказчику."),
]
w3.append(["Параметр","Значение / источник"])
for c in w3[1]: c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="305496")
for a,b in notes: w3.append([a,b])
w3.column_dimensions["A"].width=24; w3.column_dimensions["B"].width=120
for row in w3.iter_rows():
    for c in row: c.alignment=Alignment(wrap_text=True,vertical="top")
w4=wb.create_sheet("Цены на соль (источники)")
w4.append(["Источник","Производитель","Продукт / фасовка","Цена от, ₽/т","Условия","Ссылка"])
for c in w4[1]: c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="305496")
for r in sorted(REF,key=lambda x:x[3]):
    w4.append(list(r)); R=w4.max_row; w4.cell(R,4).number_format=MONEY
    if r[5].startswith("http"): w4.cell(R,6).hyperlink=r[5]; w4.cell(R,6).font=Font(color="0563C1",underline="single")
for i,w in enumerate([26,40,34,13,52,60],1): w4.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
w4.auto_filter.ref=f"A1:F{w4.max_row}"
wb.save("../SALT_STRATEGY_1410.xlsx")
for (name,*_),vs,bn,bc,pr in strat: print(name,[None if c is None else round(c) for _,c,_ in vs],bn,round(bc),round(pr))
print({k:round(v) for k,v in tot.items()},round(trec))
