"""Потенциальные тендеры (лоты 1–9): потенциальная себестоимость по ценам сайтов (розница/прайс), цель торга, маржа сейчас и после торга."""
import json, importlib.util, sys, io, re, openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side
import ref2026 as R
def load(n,f):
    s=importlib.util.spec_from_file_location(n,f); m=importlib.util.module_from_spec(s); o=sys.stdout; sys.stdout=io.StringIO(); s.loader.exec_module(m); sys.stdout=o; return m
tt=load("tt","tenderer_1_9.py"); iu=tt.iu; ss=iu.it.ft  # final_top
P=load("ssx","salt_strategy_1410.py"); Pf,be=P.P,P.be
TV=json.load(open("data/tvc2026.json"))
DISC={"розница":0.15,"прайс юрлиц":0.10,"опт":0.05}
def ptype(x):
    t=(x.get("price_type") or "").lower()
    if "розниц" in t and "опт" not in t: return "розница"
    if "опт" in t and "розниц" in t: return "прайс юрлиц"
    if "опт" in t: return "опт"
    return "прайс юрлиц"
SR={x["url"]:x for x in R.SALTROWS}
def pinfo(p):
    g=p.get("group")
    if g=="Соль":
        u=p["urls"][0] if p["urls"] else ""; x=SR.get(u) or next((y for y in R.SALTROWS if y.get("site","") in u),{})
        return ("розница" if "розниц" in (x.get("price_type") or "") else "прайс с сайта",0.10,u)
    if g in R.CLIENT: return ("ваша себестоимость",0.0,p["urls"][0] if p["urls"] else "")
    S=R.sources(g,1)
    if not S: return ("оценка",0.0,"")
    t=ptype(S[0]); return (t,DISC.get(t,0.10),S[0]["url"])
MONEY='#,##0" ₽"'; thin=Side(style="thin",color="BFBFBF")
wb=openpyxl.Workbook(); ws=wb.active; ws.title="Потенциал и торг"
ws.append(["Наши потенциальные тендеры (лоты 1–9): потенциальная себестоимость и маржа — сейчас и после торга с поставщиком"]); ws["A1"].font=Font(bold=True,size=13)
ws.append(["ВНИМАНИЕ: себестоимость — цены с сайтов (розница или прайс для юрлиц, НЕ договорной опт). На звонке проговорить торг: опт/объём/предоплата. Цель торга: розница −15%, прайс юрлиц −10%, опт −5%, соль −10%. Прибыль — чистая (УСН 15%) при ожидаемом снижении на торгах."])
H=["№","Лот / позиция","№ закупки","Подача","Кол-во","Ед.","Цена за ед. в НМЦК, ₽","Себестоимость за ед. сейчас (цена с сайта), ₽","Тип цены на сайте","Цель торга","Себестоимость за ед. после торга, ₽","Маржа сейчас","Маржа после торга","НМЦК, ₽","Себестоимость лота сейчас, ₽","Себестоимость лота после торга, ₽","Прибыль сейчас, ₽","Прибыль после торга, ₽","Где цена (ссылка)"]
ws.append(H)
for c in ws[3]: c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="305496")
LOTF=PatternFill("solid",fgColor="DDEBF7")
DSC={"Соль":.15,"ПГР":.15,"Спецодежда (коммерч":.07,"Спецодежда":.10,"Арматура":.14,"Трубы":.10}
tot=[0,0]
for k,r in enumerate(tt.ROWS,1):
    cat,n,cust,meth,end,nm,c,m,pe,b,sec,note,conf=r
    d=next((v for kk,v in DSC.items() if cat.startswith(kk)),0.10)
    if n in TV:
        pos=TV[n]["positions"]
        rows=[]; C0=C1=0; sn=sc0=sc1=0
        for p in pos:
            t,dsc,u=pinfo(p); c0=p["cost_unit"]; c1=c0*(1-dsc); q=p["qty"]; nu=p["nmck_unit"]
            rows.append((p["item"],q,p["unit"],nu,c0,t,dsc,c1,u))
            if q: C0+=c0*q; C1+=c1*q
            if nu: sn+=nu; sc0+=c0; sc1+=c1
        if n in ("32616453117",):            # без кол-ва: корзина
            C0=nm*sc0/sn; C1=nm*sc1/sn
        base_nm=nm if nm else C0/(1-0.25)   # коммерческие без НМЦК — КП с наценкой 25%
    else:  # Т Плюс, Уфа — оценка
        ps=iu.positions(r); it_,q,u,nu,cu,src=ps[0]
        lab="цена с сайта (магазин) — торг" if n=="32616449090" else "оценка — запросить КП"
        rows=[(it_,q,u,nu,cu,lab,0.10,cu*0.9,(iu.urls(src) or [""])[0])]
        C0=cu*q; C1=C0*0.9; base_nm=nm
    p0=Pf(base_nm,C0,d); p1=Pf(base_nm,C1,d); tot[0]+=p0; tot[1]+=p1
    one=len(rows)==1
    it_,q,u,nu,c0,t,dsc,c1,url=rows[0]
    ws.append([k,f"{cat}: {cust}",n,end,q if one else None,u if one else "",nu if one else None,c0 if one else None,t if one else "по позициям",dsc if one else None,c1 if one else None,
               (nu-c0)/nu if one and nu else (base_nm-C0)/base_nm,(nu-c1)/nu if one and nu else (base_nm-C1)/base_nm,nm,C0,C1,p0,p1,""])
    R_=ws.max_row; ws.cell(R_,3).number_format="@"
    if one and url: ws.cell(R_,19).value=re.sub(r"^https?://(www\.)?","",url).split("/")[0]; ws.cell(R_,19).hyperlink=url; ws.cell(R_,19).font=Font(color="0563C1",underline="single")
    for cc in ws[R_]: cc.fill=LOTF; cc.border=Border(top=thin,bottom=thin,left=thin,right=thin)
    ws.cell(R_,2).font=Font(bold=True)
    if not nm: ws.cell(R_,14).value="без НМЦК (КП +25%)"
    if not one:
        for it_,q,u,nu,c0,t,dsc,c1,url in rows:
            ws.append(["","   "+it_,"","",q,u,nu,c0,t,dsc,c1,(nu-c0)/nu if nu else None,(nu-c1)/nu if nu else None,nu*q if nu and q else None,c0*q if q else None,c1*q if q else None,None,None,""])
            R_=ws.max_row
            if url: ws.cell(R_,19).value=re.sub(r"^https?://(www\.)?","",url).split("/")[0]; ws.cell(R_,19).hyperlink=url; ws.cell(R_,19).font=Font(color="0563C1",underline="single")
            for cc in ws[R_]: cc.border=Border(top=thin,bottom=thin,left=thin,right=thin)
    # форматы
for row in ws.iter_rows(min_row=4):
    for j in (7,8,11,14,15,16,17,18):
        if isinstance(row[j-1].value,(int,float)): row[j-1].number_format=MONEY
    for j in (10,12,13):
        if isinstance(row[j-1].value,(int,float)): row[j-1].number_format="0%"
    for j in (12,13):
        v=row[j-1].value
        if isinstance(v,(int,float)): row[j-1].fill=PatternFill("solid",fgColor="C6EFCE" if v>=.4 else ("FFEB9C" if v>=.15 else "FFC7CE"))
    t=row[8].value
    if t=="розница": row[8].fill=PatternFill("solid",fgColor="FCE4D6"); row[8].font=Font(bold=True,color="C00000")
ws.append(["","ИТОГО по лотам 1–9","","","","","","","","","","","","","","",tot[0],tot[1],""])
R_=ws.max_row
for j in (17,18): ws.cell(R_,j).number_format=MONEY
for cc in ws[R_]: cc.font=Font(bold=True)
for i,w in enumerate([4,56,20,11,8,7,13,15,14,8,15,9,9,15,15,15,15,15,18],1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
ws.freeze_panes="C4"
from tidy import tidy; tidy(wb,row_h=30)
wb.save("../POTENTIAL_TORG_1510.xlsx")
print(round(tot[0]),round(tot[1]))
