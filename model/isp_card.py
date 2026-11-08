"""Карточка предприятия ISP Group LLC (для поставщиков) → ../ISP_GROUP_COMPANY_PROFILE.docx + .pdf
Данные с сайта ispgroupgc.com (08.11.2026). Поля [ ... ] — заполнить (в открытых источниках не найдены)."""
import html, docx
from docx.shared import Pt, RGBColor
from pdfutil import html_to_pdf

ROWS = [
 ("Company profile", None),
 ("Legal name", "ISP Group LLC"),
 ("Legal form / State", "Limited Liability Company, State of Florida, USA"),
 ("Florida Document No. (Sunbiz)", "[ ____________ ]"),
 ("EIN (Tax ID)", "[ __-_______ ]"),
 ("UEI (SAM.gov)", "WA46PAMTYMQ7"),
 ("CAGE code", "217F4"),
 ("Registered / office address", "16395 Biscayne Blvd, North Miami Beach, FL 33160, United States"),
 ("Phone", "+1 929 830 0038"),
 ("E-mail", "info@ispgroupgc.com"),
 ("Website", "https://ispgroupgc.com"),
 ("Authorized signatory", "[ Full name, title (Manager / Managing Member) ]"),
 ("Business activity", "Sourcing, purchasing and delivery of goods as a principal (buys in its own name): building materials, industrial equipment, tools and hardware, metals and pipe, paper and packaging, food, commodities and beverages."),
 ("Bank details", None),
 ("Beneficiary", "ISP Group LLC"),
 ("Bank name", "[ ____________ ]"),
 ("Bank address", "[ ____________ ]"),
 ("Account number", "[ ____________ ]"),
 ("ABA routing number", "[ ____________ ]"),
 ("SWIFT / BIC", "[ ____________ ]"),
 ("Currency", "USD / EUR"),
 ("Role in this transaction", None),
 ("Role", "Buyer and payer under the contract (payment from a US bank account)."),
 ("Payment terms offered", "Secured payment: escrow / blocked funds or Alibaba Trade Assurance, released after confirmation of goods at the seller's warehouse (video check or independent inspection), or payment against documents."),
]
d = docx.Document(); d.styles["Normal"].font.name = "Calibri"; d.styles["Normal"].font.size = Pt(11)
h = d.add_heading("ISP Group LLC — Company Profile", 0)
t = d.add_table(rows=0, cols=2); t.style = "Light Grid Accent 1"
for k, v in ROWS:
    r = t.add_row().cells
    if v is None:
        m = r[0].merge(r[1]); m.text = k; m.paragraphs[0].runs[0].bold = True
    else:
        r[0].text = k; r[1].text = v
        if v.startswith("[") and r[1].paragraphs[0].runs: r[1].paragraphs[0].runs[0].font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
d.add_paragraph("\nDate: ____________        Signature: ____________        Stamp/Seal: ____________")
d.save("../ISP_GROUP_COMPANY_PROFILE.docx")

e = html.escape
css = """body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:11px;color:#111;margin:20px 30px}h1{font-size:18px;color:#1f3864;margin:0 0 4px}.sub{color:#555;margin-bottom:12px}
table{border-collapse:collapse;width:100%}td{border:1px solid #c9d3e6;padding:6px 8px;vertical-align:top}td.k{width:32%;background:#f3f6fb;font-weight:bold}tr.s td{background:#305496;color:#fff;font-weight:bold}.ph{color:#c00000}.sig{margin-top:28px}"""
rows = []
for k, v in ROWS:
    if v is None: rows.append(f"<tr class=s><td colspan=2>{e(k)}</td></tr>")
    else: rows.append(f"<tr><td class=k>{e(k)}</td><td{' class=ph' if v.startswith('[') else ''}>{e(v)}</td></tr>")
page = (f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body><h1>ISP Group LLC</h1><div class=sub>Company Profile · https://ispgroupgc.com · info@ispgroupgc.com</div>"
        f"<table>{''.join(rows)}</table><div class=sig>Date: ____________ &nbsp;&nbsp;&nbsp; Signature: ____________ &nbsp;&nbsp;&nbsp; Stamp/Seal: ____________</div></body></html>")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/isp.html"
open(hp, "w").write(page); html_to_pdf(hp, "../ISP_GROUP_COMPANY_PROFILE.pdf")
print("ok")
