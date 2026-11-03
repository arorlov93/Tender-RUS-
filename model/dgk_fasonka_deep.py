"""ДГК 32616445981 «Фасонные изделия трубопроводов» — глубокая попозиционная проверка (765 поз.).
Источники: rushydro_tgk1/dgk_items_grouped.json (первичный расчёт) + dgk_deep_*.json (4 группы, перепроверка, 03.11.2026).
→ ../DGK_FASONKA_DEEP.xlsx (Итог, Позиции, Заводы для КП — без наших цен в листе заводов)."""
import json, openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from tidy import tidy

D = "data/zakupki_0211/rushydro_tgk1"
J = lambda f: json.load(open(f"{D}/{f}"))
P = {x["n"]: x for x in J("dgk_items_grouped.json")}
G = {"dgk_deep_otv_large.json": "Отводы Дн 530–1220", "dgk_deep_komp.json": "Компенсаторы", "dgk_deep_otv17375.json": "Отводы ГОСТ 17375 до Дн 530", "dgk_deep_flang_etc.json": "Фланцы, переходы, тройники, опоры"}
NEW = {}; TOT = []; MK = []; RISK = []
for f, g in G.items():
    d = J(f)
    for x in d["items"]: NEW[x["n"]] = dict(x, grp=g)
    pess = d.get("total_if_unknown_net") or d.get("total_if_vat_unknown_is_net") or d.get("total_if_vat_unknown_plus22") or d.get("total_vat_confirmed_only") or d["total"]
    TOT.append((g, len(d["items"]), d["total_prev"], d["total"], max(pess, d["total"])))
    for m in d.get("makers", []): MK.append(dict(m, grp=g))
    for r in d.get("risks", []): RISK.append((g, r if isinstance(r, str) else json.dumps(r, ensure_ascii=False)))
LOG = J("dgk_deep_flang_etc.json")["logistics"]
log_real = 8.03e6; log_pess = 11.2e6
NMC, NMC_NET = 104196798.80, 85407212
C_base = sum(t[3] for t in TOT) + log_real; C_pess = sum(t[4] for t in TOT) + log_pess
def prof(P_, C, days=150):
    pre = P_ - C - 0.008 * P_ - C * 0.2 * days / 365
    return pre - 0.15 * max(pre, 0)

thin = Side(style="thin", color="BFBFBF"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
def hdr(ws):
    for c in ws[ws.max_row]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496"); c.alignment = Alignment(wrap_text=True, vertical="top")
def put(ws, v, fmt=None):
    v = [", ".join(map(str, z)) if isinstance(z, (list, tuple)) else (json.dumps(z, ensure_ascii=False) if isinstance(z, dict) else z) for z in v]
    ws.append(v); X = ws.max_row
    for c in ws[X]: c.border = BOX; c.alignment = Alignment(wrap_text=True, vertical="top")
    for j, f in (fmt or {}).items(): ws.cell(X, j).number_format = f
    return X
wb = openpyxl.Workbook(); s = wb.active; s.title = "Итог"
s.append(["АО «ДГК» 32616445981 — фасонные изделия трубопроводов: глубокая попозиционная проверка 765 позиций (03.11.2026). Подача 11.11.2026 10:00 МСК, rushydro.roseltorg.ru."]); s["A1"].font = Font(bold=True, size=12)
s.append(["НМЦ с НДС 22 %", NMC]); s.append(["НМЦ без НДС (сравнение с нашей ценой, мы на УСН)", NMC_NET]); s.append([])
s.append(["Группа", "Позиций", "Первичный расчёт, ₽ (с НДС)", "Перепроверка — мин. подтверждённые, ₽", "Если цены без пометки НДС окажутся без НДС, ₽"]); hdr(s)
for t in TOT: put(s, list(t), {3: "#,##0", 4: "#,##0", 5: "#,##0"})
put(s, ["Доставка по 20 грузополучателям (ДФО)", "", 6160000, log_real, log_pess], {3: "#,##0", 4: "#,##0", 5: "#,##0"})
X = put(s, ["СЕБЕСТОИМОСТЬ", 765, sum(t[2] for t in TOT) + 6160000, C_base, C_pess], {3: "#,##0", 4: "#,##0", 5: "#,##0"})
for c in s[X]: c.font = Font(bold=True)
s.append([])
s.append(["Сценарий", "Цена заявки без НДС, ₽", "Себестоимость, ₽", "Маржа, ₽", "Маржа, %", "Чистая прибыль (−0,8 % ЭТП, 20 % годовых × 150 дн., УСН 15 %), ₽"]); hdr(s)
for name, Pp, C in [("По НМЦ, база", NMC_NET, C_base), ("НМЦ −10 %, база", NMC_NET * 0.9, C_base), ("По НМЦ, пессимизм", NMC_NET, C_pess), ("НМЦ −10 %, пессимизм", NMC_NET * 0.9, C_pess)]:
    put(s, [name, Pp, C, Pp - C, (Pp - C) / Pp, prof(Pp, C)], {2: "#,##0", 3: "#,##0", 4: "#,##0", 5: "0.0%", 6: "#,##0"})
s.append([])
for r in ["ВЫВОД: НЕ ИДЁМ. Даже в базовом сценарии маржа при −10 % около 6 %, чистая — в минусе; в пессимизме — убыток уже по НМЦ.",
          "Блокирующий риск: ТТ п.5.1–5.4 — в заявку сертификаты ТР ТС 032 с протоколами испытаний и полным ТУ, паспорта, заключения УЗК, обоснование безопасности на каждый вид продукции и изготовителя. 27.10 участники сообщили заказчику, что заводы компенсаторов это не дают; ответы (разъяснения №2–3, ТТ ред. 02.11) не прочитаны — ЕИС недоступна.",
          "Договор рамочный до 30.12.2027, партии ≤150 дней — больше 6 мес.; 20 адресов ДФО, в т.ч. Николаевск-на-Амуре (сезонный завоз), Советская Гавань, Нерюнгри.",
          "НДС: договор ~85 млн без НДС, вероятно, выведет нас из освобождения УСН от НДС — при НДС 5 % с выручки ещё −4,3 млн (порог проверить у бухгалтера)."]:
    s.append([r]); s.cell(s.max_row, 1).alignment = Alignment(wrap_text=True)
for i, w in enumerate([46, 18, 20, 22, 22, 26], 1): s.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

w = wb.create_sheet("Позиции (765)")
w.append(["№", "Группа", "Позиция ТТ", "Ед.", "Кол-во", "Грузополучатель", "Цена первично, ₽", "Цена перепроверка, ₽", "НДС в цене", "Сумма, ₽", "Товар / изготовитель", "Соответствие ТТ", "Источники (цена — ссылка)", "Примечание"]); hdr(w)
for n in sorted(P):
    p = P[n]; x = NEW.get(n, {}); up = x.get("unit_price")
    src = x.get("sources") or []
    src_t = "\n".join(f'{s_.get("price")} — {s_.get("url")}' for s_ in src[:3]) if src else (x.get("typesize") or p.get("url") or "")
    put(w, [n, x.get("grp", ""), p["name"], p.get("unit"), p.get("qty"), (p.get("consignee") or "")[:80], p.get("unit_price"), up, str(x.get("vat", "")), (up or 0) * float(p.get("qty") or 0) or None,
            x.get("product") or x.get("maker") or "", x.get("match") or ("оценка" if x.get("estimate") else ""), src_t, x.get("note") or ""],
        {7: "#,##0.00", 8: "#,##0.00", 10: "#,##0"})
for i, wd in enumerate([5, 22, 50, 6, 7, 30, 12, 12, 10, 13, 34, 14, 50, 40], 1): w.column_dimensions[openpyxl.utils.get_column_letter(i)].width = wd
w.freeze_panes = "D2"; w.auto_filter.ref = f"A1:N{w.max_row}"

m = wb.create_sheet("Заводы для КП")
m.append(["Группа", "Завод / поставщик", "ИНН", "Город", "Сайт", "Телефон", "E-mail", "Что закрывает", "Подтверждено в 2 источниках"]); hdr(m)
for x in MK:
    put(m, [x["grp"], x.get("name"), x.get("inn"), x.get("city"), x.get("site"), x.get("phone"), x.get("email"), str(x.get("products") or x.get("positions") or x.get("closes") or "")[:300], "да" if x.get("verified") else "нет"])
for i, wd in enumerate([24, 34, 13, 16, 24, 22, 26, 50, 10], 1): m.column_dimensions[openpyxl.utils.get_column_letter(i)].width = wd
r = wb.create_sheet("Риски")
r.append(["Группа", "Риск"]); hdr(r)
for g, t in RISK: put(r, [g, t[:800]])
r.column_dimensions["A"].width = 26; r.column_dimensions["B"].width = 120
tidy(wb, row_h=28); wb.save("../DGK_FASONKA_DEEP.xlsx")
print(round(C_base), round(C_pess), [(round((p - c) / p, 3), round(prof(p, c))) for p, c in [(NMC_NET, C_base), (NMC_NET * .9, C_base), (NMC_NET, C_pess), (NMC_NET * .9, C_pess)]])
