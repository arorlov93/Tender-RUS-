"""Наши тендеры против средней себестоимости с доставкой (одежда + соль)."""
import json, re, statistics as st, importlib.util, sys, io, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
def load(name,f):
    spec=importlib.util.spec_from_file_location(name,f); m=importlib.util.module_from_spec(spec); o=sys.stdout; sys.stdout=io.StringIO(); spec.loader.exec_module(m); sys.stdout=o; return m
cr=load("cr","cost_reference_1410.py"); ss=load("ss","salt_strategy_1410.py")
P,be=ss.P,ss.be
def base(g):
    w,rows=cr.G[g]; mids=[(a+b)/2 for _,_,a,b,_,_ in rows]; opt=[(a+b)/2 for _,_,a,b,t,_ in rows if t=="опт"]
    avg=st.mean(mids); return cr.CLIENT.get(g) or (min(st.mean(opt),avg*0.9) if opt else avg*0.9), w
SPEC={"нефть_л":(8978,2.0,"Тракт 6 731 / Спецобъединение 13 219, среднее −10%"),"нефть_з":(15953,2.6,"Тракт 11 253 / Спецобъединение 24 199, среднее −10%"),
      "сигн_л":(5501,1.3,"Спецобъединение «Рефлект-2» 6 112 −10%"),"сигн_з":(10459,2.6,"Спецобъединение «Рефлект-2» зим. 11 621 −10%"),
      "кщс":(6215,2.4,"Тракт, костюм КЩС утепл. 6 215"),"энцеф":(1112,1.0,"Фабрикант 975 / Ивсотекс 1 250"),"плащ":(1168,0.8,"Тракт 971 / Фабрикант 1 365")}
def cost(item,north=False):
    s=item.lower()+(" якут" if north else "")
    if "нефт" in s: k="нефть_з" if "зим" in s else "нефть_л"; c,w,src=SPEC[k]; return c,w,src
    if "сигнал" in s and "жилет" not in s: k="сигн_з" if ("зим" in s or "пониж" in s) else "сигн_л"; c,w,src=SPEC[k]; return c,w,src
    if "кислот" in s: return SPEC["кщс"]
    if "энц" in s: return SPEC["энцеф"]
    if "плащ" in s: return SPEC["плащ"]
    rules=[("комбинезон" in s and "зим" in s and "полу" not in s,"Комбинезон зимний"),("комбинезон" in s and "полу" not in s,"Комбинезон летний"),
           ("сварщ" in s,"Костюм сварщика"),("жилет" in s and "сигнал" in s,"Жилет сигнальный"),("жилет" in s,"Жилет утепленный"),
           ("футболк" in s,"Футболка"),("халат" in s,"Халат рабочий"),("полукомбинезон" in s or ("брюки" in s and ("зим" in s or "пониж" in s)),"Полукомбинезон / брюки зимние"),
           ("куртк" in s and ("зим" in s or "пониж" in s or "утепл" in s),"Куртка утепленная"),
           (("зим" in s or "утепл" in s or "пониж" in s) and ("якут" in s or "особ" in s),"Костюм зимний северный (IV/особый пояс)"),
           ("зим" in s or "утепл" in s or "пониж" in s,"Костюм зимний рабочий (I–III пояс)"),("спасат" in s,"Костюм летний спасателя/МЧС"),("костюм" in s,"Костюм летний рабочий")]
    for cond,g in rules:
        if cond: c,w=base(g); return c,w,f"справочник: {g}"
    return None,0,"не сопоставлено"
DLV=dict(центр=15,регион=35,север=120)
d={x["num"]:x for x in json.load(open("data/deep/out_clothes_nmck.json"))}
CL=[("0358200022926000030","Ростов, форма спасателей","19.10 09:00","регион",0.17,"44-ФЗ, цены за ед. из ЕИС"),
    ("32616431081","ФБУ ЦСМ Башкортостан, спецодежда 18 поз.","19.10 08:00","регион",0.10,"223-ФЗ, цен за ед. нет — НМЦК распределена"),
    ("32616453117","Салаватводоканал, спецодежда","21.10 09:00","регион",0.10,"223-ФЗ, договор по ед. расценкам, кол-во не задано"),
    ("32616443256","Якутия ГПС, форма (подача 15.10!)","15.10 05:00","север",0.15,"223-ФЗ, цен за ед. нет")]
MONEY='#,##0" ₽"'; thin=Side(style="thin",color="BFBFBF")
wb=openpyxl.Workbook(); S=wb.active; S.title="Сводка по лотам"
def hdr(ws,h):
    ws.append(h)
    for c in ws[ws.max_row]: c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="305496"); c.alignment=Alignment(wrap_text=True,vertical="center")
S.append(["Наши тендеры против средней себестоимости с доставкой (рынок РФ, 14.10.2026)"]); S["A1"].font=Font(bold=True,size=13)
S.append(["Себестоимость = средняя по открытым ценам фабрик/оптовиков (лист «Цены» в COST_REFERENCE_1410.xlsx) + доставка ТК; соль — среднее «завод вагоном» и «склад авто». Прибыль чистая (УСН 15%)."])
hdr(S,["Категория","Лот","№","Подача","НМЦК / база, ₽","Себестоимость с доставкой, ₽","Маржа при НМЦК","Ожид. снижение","Прибыль при ожид. снижении, ₽","Безубыток до снижения","Пометки"])
W=wb.create_sheet("Одежда — по позициям")
hdr(W,["Лот","Позиция","Кол-во","Цена за ед. в НМЦК, ₽","Средняя себест. с доставкой, ₽","Маржа позиции","Сумма НМЦК, ₽","Сумма себест., ₽","Откуда себестоимость"])
rows=[]
for num,name,end,reg,dsc,note in CL:
    x=d[num]; sn=sc=0; sn_u=sc_u=0
    for p in x["positions"]:
        c,w,src=cost(p["item"],north=(reg=="север"))
        if c is None: continue
        cd=c+w*DLV[reg]; q=p.get("qty"); nu=p.get("nmck_unit")
        if num=="32616443256": nu=None
        W.append([name,p["item"],q,nu,cd,(nu-cd)/nu if nu else None,(nu*q if nu and q else None),(cd*q if q else None),src]); R=W.max_row
        for k in (4,5,7,8): W.cell(R,k).number_format=MONEY
        W.cell(R,6).number_format="0%"
        if nu: W.cell(R,6).fill=PatternFill("solid",fgColor="C6EFCE" if (nu-cd)/nu>0.3 else ("FFEB9C" if nu>cd else "FFC7CE"))
        if q: sc+=cd*q
        if nu: sn_u+=nu; sc_u+=cd
    nm=x["nmck"]
    if num=="32616453117":   # без кол-ва: корзина по 1 ед.
        m=(sn_u-sc_u)/sn_u; C=nm*(1-m); L=0; tag="маржа по корзине (по 1 ед. каждой позиции)"
    else:
        C=sc; L=0; m=(nm-C)/nm; tag=note
    pe=P(nm,C,dsc); rows.append(("Одежда",name,num,end,nm,C,m,dsc,pe,be(nm,C),tag+("; " if tag else "")+f"доставка {reg}"))
for L_ in ss.LOTS:
    v=ss.variants(L_); real=[c for _,c,_ in v[1:3] if c]; cpt=st.mean(real); nm=L_[3]; t=L_[4]; C=cpt*t
    rows.append(("Соль",L_[0],L_[1],L_[2],nm,C,(nm-C)/nm,0.15,P(nm,C,0.15),be(nm,C),f"{t:g} т; НМЦК {nm/t:,.0f} ₽/т против себест. {cpt:,.0f} ₽/т (ваша КЗ 2 500 → прибыль {P(nm,2500*t,0.15):,.0f} ₽)".replace(","," ")))
rows.sort(key=lambda r:(0 if "Самара" not in r[1] else 1,-r[8]))
for r in rows:
    S.append(list(r)); R=S.max_row; S.cell(R,3).number_format="@"
    for k in (5,6,9): S.cell(R,k).number_format=MONEY
    for k in (7,8,10): S.cell(R,k).number_format="0%"
    S.cell(R,9).fill=PatternFill("solid",fgColor="C6EFCE" if r[8]>300000 else ("FFEB9C" if r[8]>0 else "FFC7CE"))
tot=sum(r[8] for r in rows if "Самара" not in r[1] and "15.10" not in r[3])
S.append(["","ИТОГО (подача с 16.10, без Самары)","","","","","","",tot]); R=S.max_row; S.cell(R,9).number_format=MONEY
for c in S[R]: c.font=Font(bold=True)
for i,w in enumerate([9,36,21,11,15,16,9,9,17,10,60],1): S.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
for i,w in enumerate([30,52,7,13,13,9,14,14,46],1): W.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
for ws in (S,W):
    for row in ws.iter_rows(min_row=2):
        for c in row: c.alignment=Alignment(wrap_text=True,vertical="top"); c.border=Border(top=thin,bottom=thin,left=thin,right=thin)
S.freeze_panes="C4"; W.freeze_panes="C2"; W.auto_filter.ref=f"A1:I{W.max_row}"
wb.save("../TENDERS_VS_COST_1410.xlsx")
for r in rows: print(r[0],r[1][:30],round(r[4]),round(r[5]),f"{r[6]:.0%}",round(r[8]),f"{r[9]:.0%}")
print("итого",round(tot))
