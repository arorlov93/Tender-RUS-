"""ISP Group LLC: Compliance Policy (Code of Conduct, Anti-Bribery, AML, Sanctions & Export Control, Human Rights, Data Protection, Competition)
+ схема собственности. На фирменном бланке (isp_letterhead). Для KYC Petronas и других заказчиков. → ../ISP_GROUP_COMPLIANCE_POLICY.pdf, ../ISP_GROUP_SHAREHOLDING.pdf"""
import html
from isp_letterhead import render
e = html.escape
SECTIONS = [
 ("1. Purpose and scope",
  ["This Policy sets the standards of business conduct for ISP Group LLC (the Company), its member, managers, employees, agents and any person acting on its behalf (together, Personnel). It applies to all business of the Company worldwide.",
   "Compliance with this Policy and with all applicable laws is a condition of working with or for the Company. Where local law is stricter, the stricter rule applies."]),
 ("2. Code of conduct",
  ["Personnel act honestly, fairly and with integrity, avoid conflicts of interest, keep accurate books and records, and protect confidential information of the Company and its counterparties.",
   "Any actual or potential conflict of interest must be disclosed in writing to the Manager before the transaction."]),
 ("3. Anti-bribery and corruption",
  ["The Company prohibits bribery in any form, direct or indirect, to public officials or private persons, including facilitation payments. Personnel must not offer, promise, give, request or accept any payment, gift, hospitality or other advantage to improperly influence a business decision.",
   "The Company complies with the US Foreign Corrupt Practices Act, the UK Bribery Act 2010, the Malaysian Anti-Corruption Commission Act 2009 and the anti-corruption laws of every country where it does business.",
   "Gifts and hospitality are allowed only if modest, infrequent, transparent and recorded. No payments are made in cash or to accounts of persons other than the contractual counterparty in its own name."]),
 ("4. Anti-money laundering and counter-terrorist financing",
  ["The Company does not accept or make payments that it knows or suspects to be proceeds of crime or to finance terrorism. All payments are made by bank transfer through the Company's account with a US bank, only to and from accounts held in the name of the contractual counterparty.",
   "Before the first transaction the Company verifies the identity, registration, beneficial owners and bank details of each counterparty (Know Your Counterparty). Unusual payment requests, third-party payments and cash are refused and reported to the Manager."]),
 ("5. Sanctions and export control",
  ["The Company complies with the sanctions and export control laws of the United States (including OFAC and the Export Administration Regulations), the United Nations, the European Union, the United Kingdom, Malaysia and every country where it operates.",
   "Every counterparty, its owners and the end user are screened against the OFAC SDN and sectoral lists and the UN, EU and UK consolidated lists before any quotation is issued and again before contract signature and payment. The Company does not deal, directly or indirectly, with sanctioned persons or with Countries and Territories Subject to Comprehensive Sanctions, and does not supply items subject to export licensing without the required licence.",
   "The Company does not conceal or misstate the origin, destination or end user of any goods. Screening results are recorded and kept for at least five years."]),
 ("6. Human rights and labour",
  ["The Company respects internationally recognised human rights, prohibits forced and child labour, discrimination and harassment, and expects the same from its suppliers. Personnel may raise any concern without fear of retaliation."]),
 ("7. Data protection",
  ["Personal data received from counterparties is collected only for business purposes, kept secure, used only for those purposes, not disclosed to third parties without a legal basis and deleted when no longer needed."]),
 ("8. Fair competition",
  ["The Company competes fairly and does not agree with competitors on prices, bids, markets or customers, and does not exchange competitively sensitive information. Bid rigging is strictly prohibited."]),
 ("9. Suppliers and third parties",
  ["The Company works only with manufacturers and service providers that accept these principles. Agents and distributors are appointed in writing, screened under section 5 and bound by this Policy."]),
 ("10. Reporting, training and enforcement",
  ["Concerns about any breach can be reported to the Manager at info@ispgroupgc.com, including anonymously. The Company does not retaliate against good-faith reporting.",
   "All Personnel receive this Policy on joining and confirm in writing that they have read it. Compliance training covering anti-bribery, AML, sanctions and competition is held at least once a year. Breaches lead to disciplinary action up to termination and, where required, reporting to the authorities.",
   "The Manager is responsible for this Policy, reviews it at least once a year and keeps the records of training, screening and reported concerns."]),
]
head = ("<div class=meta><div class=ref><b>ISP-POL-001 · Version 1.0</b><br>Effective 10 November 2026</div>"
        "<div class=to><div class=l>Approved by</div>Sole Member and Manager<br><b>Aleksandr Orlov</b></div></div>"
        "<h1>Compliance Policy and Code of Business Conduct</h1>")
sec = lambda S: "".join(f"<h2>{e(t)}</h2>" + "".join(f"<p>{e(p)}</p>" for p in ps) for t, ps in S)
sign = ("<div class=sign><div>Adopted by written decision of the<br>Sole Member, 10 November 2026<br><b style='color:#0B2545'>ISP Group LLC</b></div>"
        "<div><div class=ln></div><div class=cap>signature / seal</div></div><div class=nm><b>Aleksandr Orlov</b>Sole Member and Manager</div></div>")
import subprocess, os, re
from isp_letterhead import page, SCRATCH
fix = "<style>p{margin:0 0 5px} h2{margin:9px 0 4px} .page{page-break-after:always}</style>"
p1 = page(fix + head + sec(SECTIONS[:5])); p2 = page(fix + sec(SECTIONS[5:]) + sign)
b2 = re.search(r"<body>(.*)</body>", p2, re.S).group(1)
full = p1.replace("</body>", b2 + "</body>")
hp = f"{SCRATCH}/isp_policy2.html"; open(hp, "w").write(full)
subprocess.run(["/opt/pw-browsers/chromium", "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer",
                f"--print-to-pdf={os.path.abspath('../ISP_GROUP_COMPLIANCE_POLICY.pdf')}", "file://" + hp], check=True, capture_output=True, timeout=180)

sh = """<h1>Shareholding Structure</h1><p>ISP Group LLC has no parent company and no subsidiaries. The sole member and ultimate beneficial owner holds 100 % of the membership interest.</p>
<div style='margin:30px auto;width:70%;text-align:center;font-size:12px'>
<div style='border:2px solid #0B2545;padding:12px;border-radius:4px'><b>Aleksandr Orlov</b><br>Sole Member and Manager · Ultimate Beneficial Owner<br><span style='color:#6B7280'>Resident of Miami, Florida, USA</span></div>
<div style='height:34px;border-left:2px solid #B08D57;width:0;margin:0 auto'></div><div style='color:#B08D57;font-weight:700;margin:-26px 0 10px 60px;text-align:left'>100 %</div>
<div style='border:2px solid #0B2545;padding:12px;background:#F6F8FB;border-radius:4px'><b>ISP Group LLC</b><br>Wyoming LLC, incorporated 06.01.2026 · Foreign LLC in Florida, Doc. No. M26000003981 · EIN 38-4380828<br>16395 Biscayne Blvd, North Miami Beach, FL 33160, USA</div></div>
<p style='text-align:center;color:#6B7280'>No other shareholders, nominees, trusts or controlling persons.</p>
<div class=sign><div>Certified correct,<br><b style='color:#0B2545'>ISP Group LLC</b></div><div><div class=ln></div><div class=cap>signature / date</div></div><div class=nm><b>Aleksandr Orlov</b>Sole Member and Manager</div></div>"""
render(sh, "../ISP_GROUP_SHAREHOLDING.pdf", "shareholding")
print("ok")
