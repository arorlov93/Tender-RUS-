"""Разбор счёта ДН.РУ (pdftotext -layout) в JSON: строки счёта с № по заявке, кол-вом, ценой без НДС, суммой с НДС, сроком."""
import re, json, sys

f = lambda s: float(s.replace(" ", "").replace(",", "."))


def parse(txt):
    t = txt.split("поставки", 1)[1].split("Итого:")[0]
    parts = re.split(r"№(\d+) по заявке", t); out = []
    for i in range(0, len(parts) - 1, 2):
        b = parts[i]
        m = re.search(r"(\d+) шт\s+([\d ]+,\d\d) ₽\s+([\d ]+,\d\d) ₽\s+([\d ]+,\d\d) ₽", b)
        body = [l for l in b.splitlines() if l.strip() and not re.search(r"Страница|Цена в|без НДС|Наименование товара", l)]
        name = " ".join(re.split(r"\s{3,}", re.sub(r"^\s*\d+\s{2,}", "", l).strip())[0] for l in body if not re.match(r"\s*Арт\.", l))
        name = re.sub(r"\s*Ссылка.*|\s+\d+$", "", re.sub(r"\s+", " ", name)).strip()
        srok = "склад Москва" if "На складе" in b else ((re.search(r"(\d+-\d+)\s*\n?\s*раб", b) or [0, "?"])[1] + " раб.дн.")
        out.append(dict(z=int(parts[i + 1]), name=name[:220], qty=int(m[1]), unit_novat=f(m[2]), sum_vat=f(m[4]), art=(re.search(r"Арт\.: (\S+)", b) or [0, ""])[1], srok=srok))
    return out


if __name__ == "__main__":
    src, dst, title, total = sys.argv[1:5]
    items = parse(open(src).read())
    json.dump(dict(invoice=title, supplier="ООО «ДН.РУ», ИНН 7724452117", total_vat=float(total), items=items), open(dst, "w"), ensure_ascii=False, indent=1)
    print(dst, len(items), "строк", len({x["z"] for x in items}), "позиций", round(sum(x["sum_vat"] for x in items), 2))
