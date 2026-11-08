#!/usr/bin/env python3
"""Проверка компании ЕС: VIES (действительность VAT, название, адрес) + для Польши белый список МинФина
(статус НДС, KRS, REGON, дата регистрации, банковские счета).
Использование: python3 check_company.py PL 9591472409   |   python3 check_company.py PT 515968579
"""
import sys, json, re, datetime, urllib.request, subprocess


def get(url, data=None):
    req = urllib.request.Request(url, data=json.dumps(data).encode() if data else None,
                                 headers={"Content-Type": "application/json", "User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=40) as r:
        return json.load(r)


def vies(cc, num):
    try:
        d = get("https://ec.europa.eu/taxation_customs/vies/rest-api/check-vat-number", {"countryCode": cc, "vatNumber": num})
        return {"valid": d.get("valid"), "name": d.get("name"), "address": (d.get("address") or "").replace("\n", ", ")}
    except Exception as e:
        return {"error": str(e)}


def pl_whitelist(nip):
    day = datetime.date.today().isoformat()
    try:
        # API МинФина отдаёт защитную HTML-страницу для urllib — используем curl
        raw = subprocess.run(["curl", "-s", "-m", "40", f"https://wl-api.mf.gov.pl/api/search/nip/{nip}?date={day}"], capture_output=True, text=True).stdout
        if raw.lstrip().startswith("<"):
            return {"error": "API белого списка временно заблокировал запрос (защита от ботов). Проверьте вручную: https://www.podatki.gov.pl/wykaz-podatnikow-vat-wyszukiwarka (поиск по NIP, видны счета)"}
        s = json.loads(raw)["result"]["subject"]
    except Exception as e:
        return {"error": str(e)}
    if not s:
        return {"found": False}
    return {k: s.get(k) for k in ("name", "statusVat", "krs", "regon", "registrationLegalDate", "residenceAddress", "workingAddress", "accountNumbers", "removalDate")}


if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    cc, num = sys.argv[1].upper(), re.sub(r"\D", "", sys.argv[2]) if sys.argv[1].upper() != "NL" else sys.argv[2]
    out = {"country": cc, "number": num, "vies": vies(cc, num)}
    if cc == "PL":
        out["pl_white_list"] = pl_whitelist(num)
        out["note"] = "Счёт в инвойсе обязан совпадать с accountNumbers (иначе платёж > 15 тыс. PLN рискован и может быть мошенничеством)."
    print(json.dumps(out, ensure_ascii=False, indent=1))
