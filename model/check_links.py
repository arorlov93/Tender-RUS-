"""Проверка всех ссылок-источников цен: доступность, цены, контакты, признаки актуальности (2026)."""
import importlib.util, sys, io, re, subprocess, json, time, html
def load(n,f):
    s=importlib.util.spec_from_file_location(n,f); m=importlib.util.module_from_spec(s); o=sys.stdout; sys.stdout=io.StringIO(); s.loader.exec_module(m); sys.stdout=o; return m
cr=load("cr","cost_reference_1410.py"); ss=load("ss","salt_strategy_1410.py"); iu=load("iu","itog_units.py")
U={}
for g,(w,rows) in cr.G.items():
    for s_,m_,a,b,t,u in rows: U.setdefault(u,set()).add(f"одежда: {g}")
for r in ss.REF:
    if r[5].startswith("http"): U.setdefault(r[5],set()).add("соль")
for k,v in iu.SPECURL.items():
    for u in v: U.setdefault(u,set()).add(f"спец: {k}")
U.pop("https://www.trakt.ru/",None)
res=[]
PH=re.compile(r"(\+7|8)[\s\-\(]*\d{3}[\s\-\)]*\d{3}[\s\-]?\d{2}[\s\-]?\d{2}")
EM=re.compile(r"[\w.\-]+@[\w\-]+\.[a-z]{2,}",re.I)
PR=re.compile(r"\d[\d\s ]{1,8}(?:[.,]\d{1,2})?\s*(?:₽|руб|р\.)",re.I)
for u in sorted(U):
    out="/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/lk.html"
    p=subprocess.run(["curl","-sL","-m","45","-A","Mozilla/5.0 (Windows NT 10.0; Win64; x64)","-o",out,"-w","%{http_code}|%{url_effective}|%{content_type}",u],capture_output=True,text=True)
    code,eff,ct=(p.stdout.split("|")+["",""])[:3]
    txt=""
    try:
        raw=open(out,"rb").read()
        if "pdf" in ct or u.endswith(".pdf"):
            subprocess.run(["pdftotext","-l","3",out,out+".txt"],capture_output=True); txt=open(out+".txt",errors="ignore").read()
        else:
            s=raw.decode("utf-8",errors="ignore")
            if "�" in s[:5000] or "charset=windows-1251" in s.lower(): s=raw.decode("cp1251",errors="ignore")
            s=re.sub(r"<script.*?</script>|<style.*?</style>","",s,flags=re.S); txt=html.unescape(re.sub(r"<[^>]+>"," ",s))
    except Exception: pass
    years=sorted(set(re.findall(r"\b(202[3-7])\b",txt)))
    upd=re.search(r"(обновлен\w*|актуальн\w*|цены на|прайс[- ]лист от|по состоянию на)[^\n]{0,40}?(\d{1,2}[.\s]\w+[.\s]20\d\d|20\d\d)",txt,re.I)
    res.append(dict(url=u,used="; ".join(sorted(U[u])),code=code,final=eff,
        prices=len(PR.findall(txt)),phone=(PH.search(txt).group(0) if PH.search(txt) else ""),email=(EM.search(txt).group(0) if EM.search(txt) else ""),
        years=",".join(years),upd=(upd.group(0)[:70] if upd else "")))
    print(code,len(PR.findall(txt)),bool(PH.search(txt)),",".join(years),u[:80]); time.sleep(1)
json.dump(res,open("data/link_check.json","w"),ensure_ascii=False,indent=1)
