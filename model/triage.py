"""Отбор тендеров из выгрузок клиента (xlsx, 12.10.2026)."""
import openpyxl, glob, sys, re, json, datetime as dt
SRC=sys.argv[1]; OUT=sys.argv[2]
CAT={"12.10.2026.xlsx":"Соль и сыпучие","_1.xlsx":"Одежда","_2.xlsx":"Амуниция и снаряжение","_3.xlsx":"СИЗ",
     "_5.xlsx":"Вентиляция","_6.xlsx":"Сантехника и ЖКХ","_7.xlsx":"Медрасходники","_8.xlsx":"ЭОМ (кабель)","_9.xlsx":"СС"}
BAD=re.compile(r"монтаж|работ|услуг|мониторинг цен|ремонт|обслуживан|строительств|СМР|проектир|аренд|доставк[аи] .*контейнер|транспортного средства",re.I)
NOW=dt.datetime(2026,10,13,0,0)
H=None; res=[]
for f in sorted(glob.glob(SRC+"/*.xlsx")):
    cat=next(v for k,v in CAT.items() if f.endswith(k))
    rows=list(openpyxl.load_workbook(f,read_only=True,data_only=True).worksheets[0].iter_rows(values_only=True))
    H=rows[1]
    for r in rows[2:]:
        if not r or not r[0]: continue
        d=dict(zip(H,r))
        name=str(r[2] or ""); nm=d.get("НМЦ")
        try: nm=float(nm)
        except Exception: nm=None
        end=d.get("Окончание приема заявок")
        reason=[]
        if BAD.search(name): reason.append("работы/услуги")
        if not nm: reason.append("нет НМЦК")
        if isinstance(end,dt.datetime) and end<NOW: reason.append("срок прошёл")
        if nm and nm>50_000_000: reason.append(">50 млн")
        res.append(dict(cat=cat,num=str(d.get("Номер")),name=name,nmck=nm,end=str(end)[:16],law=d.get("Тип торгов"),
            method=d.get("Способ отбора"),etp=d.get("ЭТП"),smp=d.get("СМП, СОНО"),region=d.get(H[24]),cust=r[25],
            sec_app=d.get("Обеспечение заявки"),sec_contr=d.get("Обеспечение контракта"),advance=d.get("Аванс"),
            link=d.get("Ссылка на ЕИС"),ok=not reason,reason=", ".join(reason)))
json.dump(res,open(OUT,"w"),ensure_ascii=False,indent=1,default=str)
import collections
c=collections.Counter((x["cat"],x["ok"]) for x in res)
for cat in CAT.values(): print(f"{cat}: всего {c[(cat,True)]+c[(cat,False)]}, прошло фильтр {c[(cat,True)]}")
