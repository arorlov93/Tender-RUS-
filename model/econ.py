"""Сценарная экономика 15 позиций. Все суммы в рублях.
Налоговая модель по умолчанию: ООО на ОСН, НДС 22% (соль - 10%), НДС транзитный,
поэтому выручка и себестоимость считаются без НДС. Налог на прибыль 25% - отдельной строкой.
Себестоимость (cogs) - ГИПОТЕЗА в долях от цены контракта, требует подтверждения КП."""
import json, sys

RATE = 0.20          # стоимость денег, % годовых (ключевая 14% + спред)
BANK = 0.003         # РКО/эквайринг/спецсчёт, доля выручки
ETP = 0.005          # тариф ЭТП победителя / сопровождение, доля выручки
GUAR = 0.002         # банковская гарантия исполнения/заявки (~3% год. на 5-10% суммы)

# name, lot_incl_vat, vat, cogs{c,b,o}, logistics, reserve_claims, penalty_risk, cycle_days, docs_rub, deferral_days
ITEMS = [
 ("Счётчики воды Ду15-20 с поверкой",        300_000, .22, (.78,.65,.55), .02, .015, .005, 40, 0, 30),
 ("Фильтры для грузовой техники (КАМАЗ/ГАЗ/МТЗ)", 400_000, .22, (.72,.60,.50), .03, .02, .01, 40, 0, 30),
 ("Соль таблетированная NaCl 99,7% 25 кг",   1_600_000, .10, (.70,.62,.55), .12, .01, .005, 35, 0, 0),
 ("Торцевые уплотнения насосов (аналоги)",   1_800_000, .22, (.70,.55,.45), .015, .03, .01, 60, 0, 0),
 ("Фильтры карманные F7/G4 592x592",          800_000, .22, (.75,.65,.55), .05, .015, .005, 45, 0, 30),
 ("Дезсредства-концентраты (РФ, РРПП)",       500_000, .22, (.80,.65,.55), .03, .01, .005, 40, 0, 30),
 ("Химия для бассейнов (гипохлорит, pH-)",    900_000, .22, (.78,.60,.50), .06, .02, .01, 40, 0, 0),
 ("Полиоксихлорид алюминия (порошок)",       2_000_000, .22, (.85,.75,.68), .06, .015, .01, 50, 0, 0),
 ("ПГМ фасованный 25 кг (ХКНм/ПГРт)",        1_300_000, .22, (.75,.60,.50), .08, .01, .01, 40, 0, 0),
 ("Краны шаровые фланцевые КШЦФ Ду50-150",    650_000, .22, (.82,.72,.62), .03, .01, .005, 40, 0, 30),
 ("Нефтесорбенты и маты/боны",               3_400_000, .22, (.78,.65,.55), .04, .01, .01, 50, 0, 0),
 ("Подшипники ЕПК/ГПЗ (реестр ПП 719)",       500_000, .22, (.85,.75,.68), .015, .02, .005, 45, 0, 30),
 ("Вакуумные пробирки (РФ)",                  260_000, .22, (.85,.70,.55), .02, .01, .005, 40, 0, 0),
 ("Моющая химия для ПММ/пищеблоков (РФ)",     150_000, .22, (.82,.70,.60), .03, .01, .005, 40, 0, 30),
 ("Аккумуляторы AGM 12В 7-9 Ач",              250_000, .22, (.85,.72,.62), .02, .02, .005, 40, 0, 30),
]

def deal(lot, vat, cogs_ratio, log, rc, pen, days, docs, sale_adj=0.0):
    R = lot/(1+vat)*(1-sale_adj)            # выручка без НДС после снижения цены
    C = lot/(1+vat)*cogs_ratio               # закупка без НДС (фиксирована, цену продажи снижаем)
    L = R*log
    fin = (C+L)*RATE*days/365
    other = R*(BANK+ETP+GUAR) + docs
    risk = R*(rc+pen)
    op = R-C-L-fin-other-risk
    return dict(R=R,C=C,GP=R-C,GM=(R-C)/R,MU=(R-C)/C,L=L,fin=fin,other=other,risk=risk,op=op,opm=op/R,
                net=op*0.75)

def max_cogs(lot,vat,log,rc,pen,days,docs,target=0.05):
    # макс. закупка без НДС, при которой операционная маржа >= target
    R=lot/(1+vat)
    k=1+RATE*days/365
    return (R*(1-target-BANK-ETP-GUAR-rc-pen) - R*log*k - docs)/k

out=[]
for (n,lot,vat,cg,log,rc,pen,days,docs,defer) in ITEMS:
    sc={}
    for lab,ratio in zip(("cons","base","opt"),cg):
        sc[lab]=deal(lot,vat,ratio,log,rc,pen,days,docs)
    sens={d:deal(lot,vat,cg[1],log,rc,pen,days,docs,sale_adj=d)["opm"] for d in (.03,.05,.10)}
    mc=max_cogs(lot,vat,log,rc,pen,days,docs)
    b=sc["base"]
    wcA=b["C"]*(1+vat)+b["L"]*(1+vat)+lot*0.05   # закупка с НДС + логистика + обеспечение 5% деньгами
    wcB=max(0,(days-defer)/days)*b["C"]*(1+vat)+b["L"]*(1+vat)+lot*0.01  # отсрочка + гарантия
    turns=365/days
    out.append(dict(name=n,lot=lot,vat=vat,days=days,sc=sc,sens=sens,maxc=mc,maxc_ratio=mc/(lot/(1+vat)),
                    wcA=wcA,wcB=wcB,year_profit_3=b["op"]*3,roi_cycle=b["op"]/wcA))

def f(x): return f"{x:,.0f}".replace(",", " ")
for o in out:
    s=o["sc"]
    print(f"## {o['name']} | лот {f(o['lot'])} с НДС {int(o['vat']*100)}% | цикл {o['days']} дн")
    for k in ("cons","base","opt"):
        d=s[k]
        print(f"  {k}: R={f(d['R'])} C={f(d['C'])} GP={f(d['GP'])} GM={d['GM']*100:.1f}% MU={d['MU']*100:.1f}% "
              f"log={f(d['L'])} fin={f(d['fin'])} other={f(d['other'])} risk={f(d['risk'])} OP={f(d['op'])} ({d['opm']*100:.1f}%) net25={f(d['net'])}")
    print("  sens -3/-5/-10%: " + " / ".join(f"{v*100:.1f}%" for v in o['sens'].values()))
    print(f"  maxC(op>=5%) = {f(o['maxc'])} без НДС = {o['maxc_ratio']*100:.1f}% цены; WC_A={f(o['wcA'])} WC_B={f(o['wcB'])}; ROI/цикл={o['roi_cycle']*100:.1f}%")
json.dump(out,open(sys.argv[1] if len(sys.argv)>1 else "/dev/null","w"),ensure_ascii=False,indent=1,default=float)
