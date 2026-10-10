"""Сводный расчёт маржи по открытым тендерам для каждой подгруппы каталога.
Входные данные: JSON-файлы агентов (scratchpad/subg/*.json) копируются в model/data/subg/.
Формула (как в tenders.py): УСН 15%, себестоимость с НДС, прочие 0,8% цены, деньги 20% годовых за цикл."""
import json, glob, os, sys
RATE=0.20; OTHER=0.008
CATS=[("salt","1. Техническая соль и ПГМ"),("clothing","2. Одежда"),("gear","3. Амуниция и снаряжение"),
      ("ppe","4. СИЗ и бронезащита"),("vent","5. Вентиляция"),("plumb","6. Сантехника и ЖКХ"),
      ("med","7. Медицинские и лабораторные расходники"),("cable","8. ЭОМ (кабель)"),("ss","9. СС")]
def is_alt(c): return c.get("alt") or str(c.get("item","")).lower().startswith("альтернатив")
def calc(t,d):
    P=t["nmck"]*(1-d); C=sum(c["qty"]*c["unit_cost"] for c in t["cost"] if not is_alt(c)); L=t.get("log",0) or 0
    fin=(C+L)*RATE*(t.get("days") or 40)/365; pre=P-C-L-fin-P*OTHER
    return P,C+L,(P-C-L)/P if P else 0,pre-max(0,pre)*0.15
def be(t):
    lo,hi=0.0,0.99
    for _ in range(50):
        m=(lo+hi)/2
        if calc(t,m)[3]>0: lo=m
        else: hi=m
    return lo
def f(x): return f"{x:,.0f}".replace(","," ")
def coverage(t):
    pn=sum(p["qty"]*p["unit_nmck"] for p in t.get("positions",[]) if p.get("qty") and p.get("unit_nmck"))
    return pn
out=["# Открытые тендеры по подгруппам: экономика и маржа","",
"Расчёт для ООО на УСН 15% без НДС, прибыль после налога на одну сделку, без общих расходов компании. Цена контракта = НМЦК × (1 − ожидаемое снижение). Себестоимость — цены поставщиков с НДС. Прочие расходы: 0,8% цены на банк, ЭТП и гарантию; деньги в обороте — 20% годовых за цикл поставки. Ожидаемое снижение — медиана завершённых аналогичных закупок.","",
"Если себестоимость найдена не по всем позициям, маржа считается только по позициям с найденной ценой: НМЦК берётся по этим позициям, а не по всему лоту (см. колонку «Покрытие»).","",
"Пометка [Г] — в себестоимости есть гипотеза. Пометка [Р] — в себестоимости розничные цены, реальная оптовая закупка дешевле.","",
"**Решения:**",
"- «Участвовать» — прибыль при ожидаемом снижении больше 0, и сделка выдерживает снижение цены больше чем на 30%;",
"- «Осторожно» — прибыль больше 0, но запас по снижению меньше 30%;",
"- «Нет» — убыток при ожидаемом снижении.",""]
summary=[]
for key,title in CATS:
    path=f"data/subg/{key}.json"
    out.append(f"## {title}"); out.append("")
    if not os.path.exists(path): out+=["_Данные агента ещё не получены._",""]; continue
    data=json.load(open(path,encoding="utf-8"))
    out.append("| Подгруппа | Тендер | Подача | НМЦК, ₽ | Покрытие | Себестоимость, ₽ | Маржа при НМЦК | Ожид. снижение | Прибыль при ожид. снижении, ₽ | Безубыточно до | Решение |")
    out.append("|---|---|---|---|---|---|---|---|---|---|---|")
    notes=[]
    for t in data:
        if not t.get("id") or not t.get("cost"):
            out.append(f"| {t.get('subgroup','')} | — | — | — | — | — | — | — | — | — | Нет открытых тендеров / нет цен |")
            if t.get("notes"): notes.append(f"- {t.get('subgroup','')}: {t['notes']}")
            continue
        # маржа считается только по покрытым позициям
        covered_nmck=None
        try:
            names=[c.get("item","") for c in t["cost"]]
        except Exception: pass
        cov=t.get("covered_nmck")
        tt=dict(t)
        if cov: tt["nmck"]=cov
        disc=t.get("disc") or [0,0.15,0.3]
        P0,C,gm0,n0=calc(tt,0); Pb,_,gmb,nb=calc(tt,disc[1]); b=be(tt)
        flags=("[Г]" if any(c.get("hyp") for c in t["cost"]) else "")+("[Р]" if any(c.get("retail") for c in t["cost"]) else "")
        dec="**Нет**" if nb<=0 else ("Осторожно" if b<0.30 else "**Участвовать**")
        covtxt=f"{tt['nmck']/t['nmck']*100:.0f}%" if t.get("nmck") else "—"
        out.append(f"| {t['subgroup']} | [{t['id']}]({t.get('src','')}) {t.get('customer','')} ({t.get('law','')}) | {t.get('deadline','')} | {f(t['nmck'])} | {covtxt} | {f(C)} {flags} | {gm0*100:.0f}% | −{disc[1]*100:.0f}% | **{f(nb)}** | {b*100:.0f}% | {dec} |")
        summary.append((title,t['subgroup'],t['id'],nb,dec))
        if t.get("risks"): notes.append(f"- {t['id']}: {t['risks']}")
    if notes: out+=["","Риски и примечания:"]+notes
    out.append("")
out+=["## Итог по категориям","","| Категория | Тендеров с расчётом | Из них «Участвовать» | Суммарная прибыль по «Участвовать» и «Осторожно», ₽ |","|---|---|---|---|"]
for key,title in CATS:
    rows=[s for s in summary if s[0]==title]
    if not rows: continue
    out.append(f"| {title} | {len(rows)} | {sum(1 for r in rows if 'Участвовать' in r[4])} | {f(sum(r[3] for r in rows if r[3]>0))} |")
open("../SUBGROUP_TENDERS.md","w",encoding="utf-8").write("\n".join(out)+"\n")
print("ok", len(summary))
