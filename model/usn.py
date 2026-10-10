"""Сравнение ОСН (НДС 22%, прибыль 25%) и УСН 15% «доходы минус расходы» без НДС
(освобождение по ст. 145 НК: доход за прошлый год <= 20 млн руб. для 2026 г.).
Цена контракта одинаковая (на торгах побеждает одна и та же сумма), закупка у поставщика с НДС 22%.
На УСН входящий НДС не вычитается и включается в расходы (пп. 8 п. 1 ст. 346.16 НК)."""
from econ import ITEMS, RATE, BANK, ETP, GUAR

def osn(lot, vat, cg, log, rc, pen, days):
    R = lot/(1+vat); C = lot/(1+vat)*cg; L = R*log
    fin = (C+L)*RATE*days/365; other = R*(BANK+ETP+GUAR); risk = R*(rc+pen)
    op = R-C-L-fin-other-risk
    return op, op*0.75

def usn15(lot, vat, cg, log, rc, pen, days):
    R = lot                                   # вся цена контракта - доход, НДС не начисляется
    C = lot/(1+vat)*cg*(1+vat)                # та же закупка, но НДС в себестоимости
    L = lot/(1+vat)*log*(1+vat)
    fin = (C+L)*RATE*days/365; other = R*(BANK+ETP+GUAR); risk = R*(rc+pen)
    op = R-C-L-fin-other-risk
    tax = max(0.15*op, 0.01*R)
    return op, op-tax

def f(x): return f"{x:,.0f}".replace(",", " ")
print("| Позиция | Цена контракта | ОСН: после налогов | УСН 15%: после налога | Разница |")
print("|---|---|---|---|---|")
for (n,lot,vat,cg,log,rc,pen,days,docs,defer) in ITEMS:
    _,a=osn(lot,vat,cg[1],log,rc,pen,days); _,b=usn15(lot,vat,cg[1],log,rc,pen,days)
    print(f"| {n} | {f(lot)} | {f(a)} | {f(b)} | +{(b/a-1)*100:.0f}% |")
