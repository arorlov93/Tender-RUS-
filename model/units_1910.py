"""Позиции и цены за единицу для лотов из выгрузок 12.10 (gosplan)."""
import json, subprocess, time, os, sys
OUT="data/units_1910.json"
cache=json.load(open(OUT)) if os.path.exists(OUT) else {}
def get(u):
    for i in range(8):
        r=subprocess.run(["curl","-s","-m","120",u],capture_output=True,text=True).stdout
        if "rate limit" in r or not r.strip(): time.sleep(25); continue
        try: return json.loads(r)
        except Exception: time.sleep(10)
    return None
def L(x): return x if isinstance(x,list) else ([x] if x else [])
def pos44(n):
    d=get(f"https://v2test.gosplan.info/fz44/purchases/{n}/notification")
    try: po=d["source"]["notificationInfo"]["purchaseObjectsInfo"]["notDrugPurchaseObjectsInfo"]["purchaseObject"]
    except Exception: return None
    out=[]
    for x in L(po):
        q=(x.get("quantity") or {}).get("value"); unit=(x.get("OKEI") or {}).get("nationalCode","")
        out.append(dict(name=x.get("name",""),qty=float(q) if q else None,unit=unit,price=float(x["price"]) if x.get("price") else None))
    return out
def pos223(n):
    d=get(f"https://v2test.gosplan.info/fz223/purchases/{n}")
    out=[]
    def walk(o):
        if isinstance(o,dict):
            for k,v in o.items():
                if k=="lotItem":
                    for it in L(v):
                        q=it.get("qty"); out.append(dict(name=it.get("additionalInfo") or (it.get("okpd2") or {}).get("name",""),qty=float(q) if q else None,unit=(it.get("okei") or {}).get("name",""),price=None))
                else: walk(v)
        elif isinstance(o,list):
            for v in o: walk(v)
    walk(d); return out or None
nums=sys.argv[1:]
for n in nums:
    if n in cache: continue
    if n.startswith("03") or n.startswith("01") or n.startswith("08") or n.startswith("02") or n.startswith("09"): p=pos44(n)
    elif n.startswith("326"): p=pos223(n)
    else: p=None
    cache[n]=p; json.dump(cache,open(OUT,"w"),ensure_ascii=False,indent=1); time.sleep(7)
print(len(cache))
