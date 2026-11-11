"""Дистрибьюторский договор ISP Group LLC — Zaozhuang Kairui (проект для согласования), EN. → ../KAIRUI_DISTRIBUTOR_AGREEMENT_DRAFT.docx + .pdf"""
import html, os, subprocess
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

TITLE = "EXCLUSIVE DISTRIBUTION AGREEMENT"
SUB = "No. ISP-KR-2026-01 · Draft v2 for discussion · __ November 2026"
PARTIES = [
 "ZAOZHUANG KAIRUI WATER TREATMENT CO., LTD., a company organized under the laws of the People's Republic of China, No.1, Fuqian South Road, Xuecheng Chemical Industrial Park, Xuecheng District, Zaozhuang City, Shandong Province, China (the \"Manufacturer\"), represented by ____________, ____________; and",
 "ISP GROUP LLC, a limited liability company organized under the laws of the State of Wyoming, USA, registered as a foreign LLC in Florida (Doc. No. M26000003981, EIN 38-4380828), 16395 Biscayne Blvd, North Miami Beach, FL 33160, USA (the \"Distributor\"), represented by Aleksandr Orlov, Managing Member.",
]
CL = [
 ("1. Products", ["1.1 \"Products\" means the following products manufactured by the Manufacturer: ATMP (liquid and solid), HEDP (liquid and solid), PBTCA, DTPMPA, other phosphonates and phosphonate-based scale and corrosion inhibitors, and RO antiscalants, as listed in Annex 1 and updated by mutual agreement."]),
 ("2. Appointment and Territory", ["2.1 The Manufacturer appoints the Distributor as its exclusive distributor and exclusive authorized representative for the Products in the Republic of Turkmenistan (the \"Territory\"), and the Distributor accepts the appointment. The Manufacturer further appoints the Distributor as a non-exclusive distributor for the Products in the Republic of Kazakhstan, at the same prices and with the same support as the Manufacturer's other partners there.",
   "2.2 During the term the Manufacturer shall not appoint any other distributor, agent or representative for the Products in the Territory (Turkmenistan), and shall not sell the Products to customers in the Territory directly or through third parties, except through or with the written consent of the Distributor. All enquiries and tenders for the Territory received by the Manufacturer shall be referred to the Distributor.",
   "2.3 The Territory may be extended to the Republic of Kazakhstan, the Republic of Uzbekistan and other Central Asian countries by a written addendum."]),
 ("3. Tenders", ["3.1 The Distributor may bid in tenders in the Territory in its own name as the Manufacturer's exclusive authorized distributor. The Manufacturer shall issue, within 3 working days of request, tender-specific letters of authorization, certificates of analysis, technical data sheets, safety data sheets, ISO certificates and other documents required by the tender.",
   "3.2 Where a tender is open to manufacturers only, the Manufacturer shall submit the bid in its own name on the commercial terms agreed with the Distributor, and the Distributor shall act as the Manufacturer's exclusive authorized representative for that tender. In such case the Manufacturer shall pay the Distributor a representative fee equal to the difference between the contract price awarded and the Manufacturer's FOB price agreed for that tender plus documented logistics costs, within 15 days after the Manufacturer receives payment from the customer.",
   "3.3 Where a tender requires legalization of documents by an embassy, registration of the Products in a state register or third-party inspection, the Manufacturer shall provide the documents and cooperation needed; the related official fees shall be agreed per tender."]),
 ("4. Volumes", ["4.1 The Parties' agreed volume for the first contract year is 200 (two hundred) metric tons of Products purchased by the Distributor or sold under clause 3.2 in the Territory and in the Republic of Kazakhstan, in total.",
   "4.2 If the target in 4.1 is not reached in the first contract year, the Manufacturer's sole remedy shall be to convert the Distributor's appointment into a non-exclusive appointment by 30 days' written notice. No penalty or compensation shall be payable.",
   "4.3 The minimum annual volume for the second and following years shall be agreed in writing at least 60 days before the end of the first contract year, based on actual tenders and orders in the Territory."]),
 ("5. Prices and Orders", ["5.1 The Manufacturer shall supply the Products to the Distributor at prices in USD, FOB Qingdao (or EXW / FCA / CIF as agreed per order), per the Manufacturer's distributor price list in Annex 2, which shall be no less favourable than prices offered to any other buyer for the Territory.",
   "5.2 The Manufacturer shall give 30 days' notice of price changes. For tenders, the Manufacturer shall hold the price agreed for that tender for the validity period required by the tender (up to 12 months), or agree a price adjustment formula in writing before the bid.",
   "5.3 Each delivery shall be made under a purchase order or sales contract referring to this Agreement. Lead time: 20 to 30 days from order confirmation unless agreed otherwise."]),
 ("6. Payment", ["6.1 Payment: 100% by T/T against copy of the bill of lading (or railway / truck consignment note) and the shipping documents, as offered by the Manufacturer. Letter of credit or documentary collection may be used by agreement for tender contracts.",
   "6.2 All payments shall be made from the Distributor's account in the USA to the Manufacturer's company account in the Manufacturer's name only."]),
 ("7. Quality, Documents and Claims", ["7.1 Each batch shall conform to the agreed specification and be supplied with certificate of analysis, quality certificate, certificate of origin, packing list and safety data sheet; Russian-language technical documents shall be provided on request.",
   "7.2 The Manufacturer shall keep retained samples of each batch for 12 months. Quality claims may be made within 60 days after arrival at destination, supported by a third-party inspection report. Non-conforming Products shall be replaced or refunded at the Manufacturer's cost.",
   "7.3 Shelf life and warranty: as stated in the TDS, not less than 12 months from production."]),
 ("8. Distributor's Obligations", ["8.1 The Distributor shall promote the Products in the Territory, track and bid in relevant tenders, keep the Manufacturer informed of tenders and market conditions, and handle registration, translations and local formalities.",
   "8.2 During the exclusive term the Distributor shall not purchase products identical to the Products from other manufacturers for resale in the Territory, except where a customer specifies another brand by name."]),
 ("9. Compliance", ["9.1 Each Party shall comply with all applicable laws, including anti-bribery and anti-corruption laws (including the US FCPA), anti-money laundering laws, and the sanctions and export control laws of the United States, the United Nations, the European Union, the United Kingdom and the People's Republic of China.",
   "9.2 Neither Party shall sell, supply or deliver the Products, directly or indirectly, to any person subject to such sanctions, or to or for use in any country or territory subject to comprehensive sanctions. Origin, destination and end user shall be stated truthfully in all documents."]),
 ("10. Trademarks and Confidentiality", ["10.1 The Distributor may use the Manufacturer's name and product names to market the Products in the Territory. No other rights to intellectual property are granted.",
   "10.2 Each Party shall keep the other Party's prices, customers, tenders and terms confidential during the term and for 2 years after it."]),
 ("11. Term and Termination", ["11.1 This Agreement enters into force on signature and remains in force for 1 (one) year. It renews automatically for further 1-year periods unless either Party gives 60 days' written notice before expiry.",
   "11.2 Either Party may terminate by written notice if the other Party materially breaches this Agreement and does not remedy the breach within 30 days of notice. Orders and tender contracts already concluded shall be performed."]),
 ("12. Governing Law and Disputes", ["12.1 This Agreement is governed by the laws of England and Wales. Disputes shall be settled by arbitration under the rules of the Hong Kong International Arbitration Centre (HKIAC), one arbitrator, seat Hong Kong, language English.",
   "12.2 This Agreement is made in English in two originals; scanned signed copies are binding."]),
]
ANNEX = ["Annex 1. Products: ATMP 50% liquid; ATMP 95% solid; HEDP 60% liquid; HEDP 90% solid; PBTCA 50%; DTPMPA 50%; RO antiscalant KR-0100; zinc-phosphonate cooling water inhibitors; other grades by agreement.",
         "Annex 2. Distributor price list (USD/MT, FOB Qingdao): to be attached on signature."]

d = Document(); s = d.sections[0]; s.left_margin = s.right_margin = Cm(2.2); s.top_margin = s.bottom_margin = Cm(1.8)
st = d.styles["Normal"]; st.font.name = "Arial"; st.font.size = Pt(10)
def P(t, b=False, al=None, sz=None, after=5):
    p = d.add_paragraph(); r = p.add_run(t); r.bold = b
    if sz: r.font.size = Pt(sz)
    if al == "c": p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if al == "j": p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(after)
P(TITLE, b=True, al="c", sz=14, after=2); P(SUB, al="c", sz=9, after=12)
for x in PARTIES: P(x, al="j")
P("together the \"Parties\", have agreed as follows:", after=8)
for t, ps in CL:
    P(t, b=True, after=3)
    for x in ps: P(x, al="j", after=3)
for x in ANNEX: P(x, al="j", sz=9, after=3)
P("", after=10)
t = d.add_table(rows=4, cols=2); t.style = "Table Grid"
for i, (a, b2) in enumerate([("MANUFACTURER", "DISTRIBUTOR"), ("Zaozhuang Kairui Water Treatment Co., Ltd.", "ISP Group LLC"), ("Name / Title: ______________", "Aleksandr Orlov, Managing Member"), ("Signature / Stamp: ______________", "Signature: ______________")]):
    t.cell(i, 0).text = a; t.cell(i, 1).text = b2
d.save("../KAIRUI_DISTRIBUTOR_AGREEMENT_DRAFT.docx")

e = html.escape
css = """@page{size:A4;margin:16mm 18mm}body{font-family:'DejaVu Sans',Arial,sans-serif;font-size:9.3px;color:#111;line-height:1.38}h1{text-align:center;font-size:14px;margin:0}.sub{text-align:center;color:#555;font-size:8.5px;margin:3px 0 10px}
h2{font-size:10px;margin:8px 0 2px}p{margin:0 0 4px;text-align:justify}table{width:100%;border-collapse:collapse;margin-top:14px}td{border:1px solid #999;padding:6px;width:50%;vertical-align:top}.ann{font-size:8.5px;color:#333}"""
h = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body><h1>{TITLE}</h1><div class=sub>{e(SUB)}</div>"]
h += [f"<p>{e(x)}</p>" for x in PARTIES] + ["<p>together the \"Parties\", have agreed as follows:</p>"]
for t2, ps in CL: h.append(f"<h2>{e(t2)}</h2>" + "".join(f"<p>{e(x)}</p>" for x in ps))
h += [f"<p class=ann>{e(x)}</p>" for x in ANNEX]
h.append("<table><tr><td><b>MANUFACTURER</b><br>Zaozhuang Kairui Water Treatment Co., Ltd.<br><br>Name / Title: ______________<br><br>Signature / Stamp: ______________</td><td><b>DISTRIBUTOR</b><br>ISP Group LLC<br><br>Aleksandr Orlov, Managing Member<br><br>Signature: ______________</td></tr></table></body></html>")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/kr_agr.html"; open(hp, "w").write("".join(h))
subprocess.run(["/opt/pw-browsers/chromium", "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer", f"--print-to-pdf={os.path.abspath('../KAIRUI_DISTRIBUTOR_AGREEMENT_DRAFT.pdf')}", "file://" + hp], check=True, capture_output=True, timeout=180)
print("ok")
