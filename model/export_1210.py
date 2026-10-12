"""Итоговая таблица по выгрузкам клиента 12.10.2026 -> TENDERS_1210.xlsx"""
import json, os, glob, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from subgroups import calc, be
SUB="/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/subg2"
tri=json.load(open("data/triage_1210.json"))
rows=[]
def P(nm,c,d):
    p=nm*(1-d); pre=p-c-p*0.008-c*0.2*60/365; return pre-max(pre,0)*0.15
# соль
for n,cust,t,nm,end,note in [("32616451106","ООО «Департамент городского хозяйства», МО",400,6724000,"16.10.2026 10:00","223-ФЗ; обеспечение исполнения 10%"),
    ("32616436306","ООО «ИКС Орехово-Зуево», МО",948,13869240,"13.10.2026 10:00","223-ФЗ; «Концентрат минеральный Галит» — сверить ТЗ на эквивалент"),
    ("32616447519","АО «Волоколамское ПТП РЖКХ», МО",130,1593800,"16.10.2026 10:00","223-ФЗ; «Концентрат минеральный Галит» — сверить ТЗ")]:
    c=t*2500; rows.append(["Соль",n,cust,f"Соль техническая {t} т ({nm/t:,.0f} ₽/т)".replace(","," "),"223-ФЗ",end,nm,c,(nm-c)/nm,0.03,P(nm,c,.03),P(nm,c,.15),"Участвовать",note])
# одежда по ценам клиента
xs=[x for x in tri if x['cat']=='Одежда' and x['ok'] and 100000<=(x['nmck'] or 0)<=20e6]; xs.sort(key=lambda x:-x['nmck'])
for x in xs:
    nm=x['nmck']; c=0.75*nm
    rows.append(["Одежда",x['num'],str(x['cust']),x['name'],x['law'],x['end'],nm,c,0.25,0.05,P(nm,c,.05),P(nm,c,.05),"Участвовать, если ≤2 заявок","себестоимость = 75% НМЦК (данные клиента), снижение ≤5%"])
# агентские категории
names={"med":"Медрасходники","cable":"Кабель (ЭОМ)","ss":"СС","gear":"Амуниция","plumb":"Сантехника и ЖКХ"}
for k,cat in names.items():
    p=f"{SUB}/{k}.json"
    if not os.path.exists(p): continue
    for t in json.load(open(p,encoding="utf-8")):
        if not t.get("id") or not t.get("cost") or not t.get("nmck"): continue
        tt=dict(t)
        if t.get("covered_nmck"): tt["nmck"]=t["covered_nmck"]
        disc=t.get("disc") or [0,.15,.3]
        P0,C,gm,n0=calc(tt,0); Pb,_,_,nb=calc(tt,disc[1]); b=be(tt)
        dec="Нет" if nb<=0 else ("Осторожно" if b<0.3 else "Участвовать")
        rows.append([cat,t["id"],t.get("customer",""),t.get("subgroup",""),t.get("law",""),t.get("deadline",""),t["nmck"],C,gm,disc[1],n0,nb,dec,(t.get("risks") or "")[:300]])
wb=openpyxl.Workbook(); ws=wb.active; ws.title="Тендеры 12.10"
H=["Категория","№ закупки","Заказчик","Предмет","Закон","Подача","НМЦК, ₽","Себестоимость, ₽","Валовая маржа при НМЦК","Ожид. снижение","Прибыль при НМЦК, ₽","Прибыль при ожид. снижении, ₽","Решение","Комментарий"]
ws.append(H)
for c in ws[1]: c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="305496"); c.alignment=Alignment(wrap_text=True,vertical="center")
order={"Участвовать":0,"Участвовать, если ≤2 заявок":1,"Осторожно":2,"Нет":3}
rows.sort(key=lambda r:(order.get(r[12],9),-r[11]))
for r in rows: ws.append(r)
for row in ws.iter_rows(min_row=2):
    for i in (6,7,10,11): row[i].number_format='# ##0'
    for i in (8,9): row[i].number_format='0%'
    col={"Участвовать":"C6EFCE","Участвовать, если ≤2 заявок":"FFEB9C","Осторожно":"FFEB9C","Нет":"FFC7CE"}.get(row[12].value)
    if col: row[12].fill=PatternFill("solid",fgColor=col)
for i,w in enumerate([16,22,40,45,10,18,14,14,10,10,14,16,18,60],1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
ws.freeze_panes="C2"; ws.auto_filter.ref=ws.dimensions
# итоги
s=wb.create_sheet("Итог")
s.append(["Категория","Лотов","«Участвовать»","Прибыль по «Участвовать», ₽"])
cats={}
for r in rows:
    c=cats.setdefault(r[0],[0,0,0]); c[0]+=1
    if r[12].startswith("Участвовать"): c[1]+=1; c[2]+=r[11]
for k,v in cats.items(): s.append([k]+v)
for c in s[1]: c.font=Font(bold=True)
s.column_dimensions["A"].width=24; s.column_dimensions["D"].width=26
for row in s.iter_rows(min_row=2): row[3].number_format='# ##0'
wb.save("../TENDERS_1210.xlsx"); print(len(rows), {k:v for k,v in cats.items()})
