"""Фирменный бланк ISP Group LLC (Letter, Inter). render(body_html, out_pdf) — письмо на бланке; при запуске — пустой бланк PDF + DOCX.
Реквизиты — как на ispgroupgc.com и в Sunbiz; банковских реквизитов здесь нет и не должно быть."""
import os, subprocess
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NAVY, GOLD, INK, MUTED = "#0B2545", "#B08D57", "#1C1C1C", "#6B7280"
CO = dict(name="ISP GROUP LLC", tag="Global Commerce", addr1="16395 Biscayne Blvd", addr2="North Miami Beach, FL 33160, USA",
          phone="+1 929 830 0038", mail="info@ispgroupgc.com", web="ispgroupgc.com", ein="38-4380828", doc="M26000003981")
SCRATCH = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad"

FONT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")  # Inter (OFL), сконвертирован в TTF: иначе Chromium встраивает Type 3
FACES = "".join(f"@font-face{{font-family:'InterTT';font-weight:{w};src:url('file://{FONT_DIR}/Inter-{n}.ttf')}}"
                for w, n in ((400, "Regular"), (500, "Medium"), (600, "SemiBold"), (700, "Bold"), (800, "ExtraBold")))
CSS = FACES + f"""
@page {{ size: Letter; margin: 0 }}
* {{ box-sizing: border-box }}
html, body {{ margin: 0; padding: 0; background: #fff }}
body {{ font-family: 'InterTT', 'DejaVu Sans', sans-serif; color: {INK}; font-size: 11.4px; line-height: 1.58;
       -webkit-print-color-adjust: exact; print-color-adjust: exact }}
.page {{ width: 8.5in; height: 11in; position: relative; padding: 0.55in 0.75in 0.9in 0.75in; overflow: hidden }}
.band {{ position: absolute; top: 0; left: 0; right: 0; height: 6px; background: {NAVY} }}
.band::after {{ content: ""; position: absolute; left: 0; bottom: -2px; width: 1.9in; height: 2px; background: {GOLD} }}
header {{ display: flex; justify-content: space-between; align-items: flex-start; padding-top: 2px }}
.mark {{ display: flex; align-items: center; gap: 12px }}
.isp {{ font-weight: 800; font-size: 30px; letter-spacing: -0.5px; color: {NAVY}; line-height: 1 }}
.bar {{ width: 1.2px; height: 34px; background: {NAVY} }}
.grp {{ font-weight: 400; font-size: 11.5px; letter-spacing: 6px; color: {NAVY}; line-height: 1 }}
.tag {{ font-size: 8.4px; letter-spacing: 2.2px; text-transform: uppercase; color: {GOLD}; margin-top: 5px; font-weight: 600 }}
.co {{ text-align: right; font-size: 8.4px; line-height: 1.55; color: {MUTED} }}
.co b {{ color: {NAVY}; font-size: 9px; letter-spacing: 1.2px; font-weight: 700 }}
.rule {{ height: 1px; background: #D9DEE5; margin: 16px 0 18px }}
footer {{ position: absolute; left: 0.75in; right: 0.75in; bottom: 0.42in; border-top: 1px solid #D9DEE5; padding-top: 7px;
         display: flex; justify-content: space-between; font-size: 7.6px; color: {MUTED}; letter-spacing: .2px }}
footer .w {{ color: {NAVY}; font-weight: 600; letter-spacing: .8px }}
/* тело письма */
.meta {{ display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 14px }}
.ref {{ font-size: 9px; color: {MUTED}; line-height: 1.6 }} .ref b {{ color: {INK}; font-weight: 600 }}
.to {{ text-align: right; font-size: 10.2px; line-height: 1.5 }} .to .l {{ font-size: 8px; color: {MUTED}; letter-spacing: 1.5px; text-transform: uppercase }}
h1 {{ font-size: 13.5px; font-weight: 700; color: {NAVY}; margin: 6px 0 14px; padding-left: 10px; border-left: 3px solid {GOLD}; line-height: 1.35 }}
p {{ margin: 0 0 8px; text-align: justify }}
.greet {{ font-weight: 600 }}
table {{ width: 100%; border-collapse: collapse; margin: 6px 0 12px; font-size: 10.4px }}
th {{ background: {NAVY}; color: #fff; text-align: left; font-weight: 600; padding: 6px 8px; font-size: 9px; letter-spacing: .3px }}
td {{ padding: 6px 8px; vertical-align: top; border-bottom: 1px solid #E5E8EE }}
tr:nth-child(even) td {{ background: #F6F8FB }}
td.g {{ font-weight: 600; color: {NAVY} }}
h2 {{ font-size: 10px; font-weight: 700; color: {NAVY}; text-transform: uppercase; letter-spacing: 1.2px; margin: 12px 0 6px }}
ul {{ margin: 0 0 10px; padding-left: 16px }} li {{ margin-bottom: 3px }} li::marker {{ color: {GOLD} }}
.sign {{ margin-top: 40px; display: grid; grid-template-columns: 1fr 1.6in 1fr; align-items: end; gap: 14px }}
.sign .ln {{ border-bottom: 1px solid {INK}; height: 34px }}
.sign .cap {{ font-size: 7.6px; color: {MUTED}; text-align: center; margin-top: 3px }}
.sign .nm {{ text-align: right }} .sign .nm b {{ display: block; color: {NAVY} }}
"""

def page(body):
    return f"""<!DOCTYPE html><html><head><meta charset='utf-8'><style>{CSS}</style></head><body><div class=page><div class=band></div>
<header><div><div class=mark><span class=isp>ISP</span><span class=bar></span><span class=grp>GROUP</span></div><div class=tag>{CO['tag']}</div></div>
<div class=co><b>{CO['name']}</b><br>{CO['addr1']}<br>{CO['addr2']}<br>{CO['phone']} · {CO['mail']}</div></header><div class=rule></div>
{body}
<footer><span>ISP Group LLC · Florida Doc. No. {CO['doc']} · EIN {CO['ein']}</span><span class=w>{CO['web'].upper()}</span></footer></div></body></html>"""

def render(body, out_pdf, name="letter"):
    hp = f"{SCRATCH}/isp_{name}.html"; open(hp, "w").write(page(body))
    subprocess.run(["/opt/pw-browsers/chromium", "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer",
                    f"--print-to-pdf={os.path.abspath(out_pdf)}", "file://" + hp], check=True, capture_output=True, timeout=180)

def blank_docx(out):
    d = Document(); s = d.sections[0]
    s.page_width, s.page_height = Cm(21.59), Cm(27.94)
    s.left_margin = s.right_margin = Cm(1.9); s.top_margin = Cm(3.4); s.bottom_margin = Cm(2.2); s.header_distance = Cm(1.0); s.footer_distance = Cm(0.9)
    st = d.styles["Normal"]; st.font.name = "Inter"; st.font.size = Pt(10); st.element.rPr.rFonts.set(qn("w:eastAsia"), "Inter")
    hdr = s.header; t = hdr.add_table(rows=1, cols=2, width=Cm(17.8))
    a = t.cell(0, 0).paragraphs[0]
    for txt, b, sz, sp in (("ISP", True, 22, 0), ("  |  ", False, 18, 0), ("G R O U P", False, 10, 0)):
        r = a.add_run(txt); r.bold = b; r.font.size = Pt(sz); r.font.color.rgb = RGBColor.from_string(NAVY[1:])
    a2 = t.cell(0, 0).add_paragraph(); r = a2.add_run("GLOBAL COMMERCE"); r.font.size = Pt(7); r.bold = True; r.font.color.rgb = RGBColor.from_string(GOLD[1:])
    c = t.cell(0, 1); c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = c.paragraphs[0].add_run(CO["name"]); r.bold = True; r.font.size = Pt(7.5); r.font.color.rgb = RGBColor.from_string(NAVY[1:])
    for line in (CO["addr1"], CO["addr2"], f"{CO['phone']} · {CO['mail']}"):
        p = c.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.RIGHT; p.paragraph_format.space_after = Pt(0)
        r = p.add_run(line); r.font.size = Pt(7); r.font.color.rgb = RGBColor.from_string(MUTED[1:])
    hp = hdr.add_paragraph(); pPr = hp._p.get_or_add_pPr(); bd = OxmlElement("w:pBdr"); bt = OxmlElement("w:bottom")
    for k, v in (("w:val", "single"), ("w:sz", "6"), ("w:space", "1"), ("w:color", GOLD[1:])): bt.set(qn(k), v)
    bd.append(bt); pPr.append(bd)
    fp = s.footer.paragraphs[0]; r = fp.add_run(f"ISP Group LLC · Florida Doc. No. {CO['doc']} · EIN {CO['ein']}        {CO['web'].upper()}")
    r.font.size = Pt(7); r.font.color.rgb = RGBColor.from_string(MUTED[1:])
    d.add_paragraph("Ref. No. ______    Date: ______"); d.add_paragraph("To: ______"); d.add_paragraph("Subject: ______")
    d.save(out)

if __name__ == "__main__":
    render("<div style='height:7.3in'></div>", "../ISP_GROUP_BLANK.pdf", "blank")
    blank_docx("../ISP_GROUP_BLANK.docx"); print("ok")
