"""Отбор лотов из выгрузок 19.10.2026: без аукционов, подача с 20.10, без работ/услуг."""
import openpyxl, glob, re, json, datetime as dt, sys, collections
SRC=sys.argv[1]
CAT={"_3.xlsx":"Одежда/СИЗ","_4.xlsx":"Соль","_5.xlsx":"Форма/снаряжение","_6.xlsx":"СИЗ (респираторы)","_7.xlsx":"Вентиляция/фильтры",
     "_8.xlsx":"Сантехника/арматура","_9.xlsx":"Медрасходники (пробирки)","_10.xlsx":"СС/видеонаблюдение","_11.xlsx":"Кабель (ЭОМ)"}
BAD=re.compile(r"монтаж|работ|услуг|ремонт|обслуживан|строительств|СМР|проектир|аренд|мониторинг цен",re.I)
AUC=re.compile(r"аукцион",re.I)
FROM=dt.datetime(2026,10,20)
res=[];seen=set()
for f in sorted(glob.glob(SRC+"/*.xlsx")):
    cat=next(v for k,v in CAT.items() if f.endswith(k))
    rows=list(openpyxl.load_workbook(f,read_only=True,data_only=True).worksheets[0].iter_rows(values_only=True))
    H=rows[1]
    for r in rows[2:]:
        if not r or not r[0]: continue
        d=dict(zip(H,r)); name=str(r[2] or ""); num=str(r[0])
        key=(num,name)
        if key in seen: continue
        seen.add(key)
        try: nm=float(d.get("НМЦ"))
        except Exception: nm=None
        end=d.get("Окончание приема заявок"); meth=str(d.get("Способ отбора") or ""); law=d.get("Тип торгов")
        why=[]
        if AUC.search(meth): why.append("аукцион")
        if BAD.search(name): why.append("работы/услуги")
        if isinstance(end,dt.datetime) and end<FROM: why.append("срок до 20.10")
        if end is None: why.append("срок не указан")
        res.append(dict(cat=cat,num=num,name=name,nmck=nm,end=str(end)[:16],law=law,method=meth,etp=d.get("ЭТП"),smp=d.get("СМП, СОНО"),
            region=r[24],cust=r[25],sec_app=d.get("Обеспечение заявки"),sec_contr=d.get("Обеспечение контракта"),advance=d.get("Аванс"),
            place=r[28] if len(r)>28 else None,ok=not why,why=", ".join(why)))
json.dump(res,open("data/triage_1910.json","w"),ensure_ascii=False,indent=1,default=str)
c=collections.Counter((x["cat"],x["ok"]) for x in res)
for cat in CAT.values(): print(f"{cat}: всего {c[(cat,True)]+c[(cat,False)]}, прошло {c[(cat,True)]}")
print(collections.Counter(x["why"] for x in res if not x["ok"]).most_common(6))
