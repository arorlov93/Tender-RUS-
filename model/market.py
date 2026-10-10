"""Объём рынка 44-ФЗ по кодам ОКПД2: сумма НМЦК опубликованных закупок (gosplan, лимит 10 запросов/мин)."""
import json, subprocess, time, sys, urllib.parse
LAW=sys.argv[2] if len(sys.argv)>2 else "fz44"
B=f"https://v2test.gosplan.info/{LAW}/purchases"
def get(p):
    u=B+"?"+urllib.parse.urlencode(p)
    for i in range(10):
        r=subprocess.run(["curl","-s","-m","120",u],capture_output=True,text=True).stdout
        if "rate limit" in r or not r.strip(): time.sleep(25); continue
        try: return json.loads(r)
        except Exception: time.sleep(10)
    return []
CODES=[("Соль","08.93.10.115","2025-08-01","2026-04-01",1.0),("Соль","08.93.10.190","2025-08-01","2026-04-01",1.0),
       ("ПГМ","20.59.43.130","2025-08-01","2026-04-01",1.0),
       ("Одежда","14.12.11","2026-07-10","2026-10-10",4),("Одежда","14.12.21","2026-07-10","2026-10-10",4),
       ("Одежда","14.12.30","2026-07-10","2026-10-10",4),("Одежда","14.14","2026-07-10","2026-10-10",4),
       ("Одежда","15.20.31","2026-07-10","2026-10-10",4),("Одежда","15.20.32","2026-07-10","2026-10-10",4),
       ("Амуниция","13.92.24","2026-07-10","2026-10-10",4),("Амуниция","15.12.12","2026-07-10","2026-10-10",4),
       ("Амуниция","13.92.22.152","2026-07-10","2026-10-10",4),
       ("СИЗ","32.99.11","2026-07-10","2026-10-10",4),("СИЗ","32.50.42","2026-07-10","2026-10-10",4),
       ("Вентиляция","28.25.14","2026-07-10","2026-10-10",4),
       ("Сантехника/ЖКХ","28.14.12","2026-07-10","2026-10-10",4),("Сантехника/ЖКХ","28.14.13","2026-07-10","2026-10-10",4),
       ("Сантехника/ЖКХ","22.21.21","2026-07-10","2026-10-10",4),("Сантехника/ЖКХ","26.51.52","2026-07-10","2026-10-10",4),
       ("Сантехника/ЖКХ","26.51.63.120","2026-07-10","2026-10-10",4),
       ("Медрасходники","32.50.50.000-00001454","2026-07-10","2026-10-10",4),
       ("Кабель (ЭОМ)","27.32","2026-07-10","2026-10-10",4),
       ("СС","26.40.33","2026-07-10","2026-10-10",4),("СС","26.30.11","2026-07-10","2026-10-10",4),("СС","26.30.50","2026-07-10","2026-10-10",4)]
res=[]
ONLY=sys.argv[3].split(",") if len(sys.argv)>3 else None
for cat,code,a,b,k in CODES:
    if ONLY and code not in ONLY: continue
    seen={}; skip=0
    while True:
        d=get({"classifier":code,"published_after":a,"published_before":b,"limit":100,"skip":skip})
        if not isinstance(d,list) or not d: break
        for x in d: seen[x["purchase_number"]]=(x.get("max_price") or 0, x.get("customer") or x.get("customer_inn") or x.get("placer"))
        if len(d)<100 or skip>=3000: break
        skip+=100; time.sleep(7)
    s=sum(v[0] for v in seen.values()); n=len(seen); c=len({v[1] for v in seen.values()})
    res.append(dict(cat=cat,code=code,window=f"{a}..{b}",n=n,cust=c,sum=s,year=s*k,capped=skip>=3000))
    print(cat,code,n,c,round(s/1e6,1),"млн; в год ≈",round(s*k/1e6,1),"млн",flush=True)
    json.dump(res,open(sys.argv[1],"w"),ensure_ascii=False,indent=1)
    time.sleep(7)
