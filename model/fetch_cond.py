"""Сырые извещения gosplan по лотам с подачей с 25.10 (для условий: аванс, обеспечение, оплата)."""
import json, re, subprocess, time, os, sys
OUT = sys.argv[1]
nums = [l.split('|')[1].strip() for l in open(sys.argv[2]) if l.count('|') > 4]
for n in nums:
    if re.fullmatch(r"0\d{18}", n): urls = {"p": f"https://v2test.gosplan.info/fz44/purchases/{n}", "n": f"https://v2test.gosplan.info/fz44/purchases/{n}/notification"}
    elif re.fullmatch(r"3\d{10}", n): urls = {"p": f"https://v2test.gosplan.info/fz223/purchases/{n}"}
    else: continue
    for k, u in urls.items():
        f = f"{OUT}/{n}_{k}.json"
        if os.path.exists(f) and os.path.getsize(f) > 200: continue
        for _ in range(6):
            r = subprocess.run(["curl", "-s", "-m", "40", u], capture_output=True, text=True).stdout
            if "rate" in r[:200].lower() or not r.strip(): time.sleep(25); continue
            open(f, "w").write(r); break
    print(n, "ok", flush=True)
