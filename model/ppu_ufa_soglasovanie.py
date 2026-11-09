"""Письмо в МУП «Уфимские инженерные сети»: согласование ГОСТ труб для отводов 377х8 и 426х8 в ППУ-ПЭ (закупка 32616450774).
Размеры, стенка, сталь — строго по ТЗ (data/ppu_ufa/tz_verbatim.json); основание — спецификация АО «ТИК» 15251.01 от 09.11.2026.
→ ../PPU_UFA_SOGLASOVANIE_GOST.docx + .pdf"""
import html, os, subprocess
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

CO = dict(name="ООО «ГИДРОПАС»", inn="9729231965", kpp="772601001",
          addr="117105, г. Москва, Варшавское шоссе, д. 1А, помещ. 28/2", tel="+7 (925) 878-06-78", mail="info@gidropas.ru")
TO = ["МУП «Уфимские инженерные сети»", "Отдел закупок, М. А. Даниловой", "zakupki@mup-uis.ru"]
SUBJ = "О согласовании ГОСТ труб для отводов 377х8 и 426х8"
INTRO = [
 "Добрый день!",
 "По договору на поставку отводов в ППУ-ПЭ (закупка № 32616450774) у завода нет в наличии труб по ГОСТ 20295-85 для отводов 377х8 и 426х8. "
 "Просим согласовать замену стандарта трубы. Размеры, толщина стенки и сталь остаются как в ТЗ.",
]
T1H = ["Позиция", "По ТЗ", "Предлагаем", "Почему не хуже"]
T1 = [
 ["Отвод 377х8-90-1-ППУ-ПЭ, 2 шт", "труба 377х8, Ст20, ГОСТ 20295-85", "труба 377х8, Ст20, ГОСТ 8731-74 (бесшовная)", "Лучше, чем труба со сварным соединением: бесшовная, сварного шва нет"],
 ["Отвод 426х8-90-1-ППУ-ПЭ, 3 шт", "труба 426х8, Ст20, ГОСТ 20295-85", "труба 426х8, Ст20, ГОСТ 10705-80, группа В", "Та же технология (прямой шов) и та же сталь. ГОСТ 10705 уже есть в вашем ТЗ для отвода 108х4,5"],
]
T2H, T2, WHY = None, None, []
KEEP = [
 "Отводы 108х4,5 и 530х9 поставляем строго по ТЗ.",
 "Цена, сроки, доставка и все остальные требования ТЗ (ГОСТ 30732-2020, ОДК, концы 150/210 мм, сертификаты) не меняются.",
]
ASK = "Просим письменно согласовать замену до 11.11.2026, чтобы запустить изделия в производство и поставить их в срок."

# ---------------- DOCX ----------------
d = Document()
for s in d.sections: s.left_margin = s.right_margin = Cm(2); s.top_margin = s.bottom_margin = Cm(1.5)
st = d.styles["Normal"]; st.font.name = "Times New Roman"; st.font.size = Pt(11)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
def P(t="", b=False, al=None, sz=None, color=None, after=4):
    p = d.add_paragraph(); r = p.add_run(t); r.bold = b
    if sz: r.font.size = Pt(sz)
    if color: r.font.color.rgb = RGBColor.from_string(color)
    if al: p.alignment = {"c": WD_ALIGN_PARAGRAPH.CENTER, "r": WD_ALIGN_PARAGRAPH.RIGHT, "j": WD_ALIGN_PARAGRAPH.JUSTIFY}[al]
    p.paragraph_format.space_after = Pt(after); return p
def shade(cell, hexc):
    tcPr = cell._tc.get_or_add_tcPr(); sh = OxmlElement("w:shd"); sh.set(qn("w:val"), "clear"); sh.set(qn("w:fill"), hexc); tcPr.append(sh)
def table(head, rows, widths, hl_col=None):
    t = d.add_table(rows=1, cols=len(head)); t.style = "Table Grid"
    for j, h in enumerate(head):
        c = t.rows[0].cells[j]; c.text = ""; r = c.paragraphs[0].add_run(h); r.bold = True; r.font.size = Pt(9.5); r.font.color.rgb = RGBColor(255, 255, 255); shade(c, "1F4E79")
    for row in rows:
        cs = t.add_row().cells
        for j, v in enumerate(row):
            cs[j].text = ""; r = cs[j].paragraphs[0].add_run(v); r.font.size = Pt(9.5)
            if hl_col is not None and j in hl_col: shade(cs[j], "E2EFDA")
    for row in t.rows:
        for j, w in enumerate(widths): row.cells[j].width = Cm(w)
    d.add_paragraph().paragraph_format.space_after = Pt(2)

P(CO["name"], b=True, sz=16, color="1F4E79", after=0)
P(f"ИНН {CO['inn']} / КПП {CO['kpp']}  ·  {CO['addr']}  ·  {CO['tel']}  ·  {CO['mail']}", sz=8.5, color="555555", after=8)
P("Исх. № ______ от 09.11.2026", sz=10, after=2)
for x in TO: P(x, al="r", after=0)
P("", after=6)
P(SUBJ, b=True, al="c", sz=12.5, after=8)
for i, x in enumerate(INTRO): P(x, al="j" if i else None)
table(T1H, T1, [3.6, 3.8, 4.4, 5.2], hl_col={2})
for x in KEEP: P(x, al="j", after=3)
P(ASK, b=True, al="j", after=10)
P("Приложение: спецификация АО «Тепло-Изоляционная Компания» № 15251.01 от 09.11.2026.", sz=10, after=14)
P("С уважением,", after=0)
P("Генеральный директор ООО «ГИДРОПАС»          ______________          /________________/", after=0)
P("М.П.", sz=9)
d.save("../PPU_UFA_SOGLASOVANIE_GOST.docx")

# ---------------- PDF (A4 портрет) ----------------
e = html.escape
css = """@page{size:A4 portrait;margin:14mm 16mm}body{font-family:'Times New Roman','DejaVu Serif',serif;font-size:12.5px;color:#111}
.co{font-size:20px;font-weight:bold;color:#1F4E79}.rk{font-size:8.5px;color:#555;border-bottom:2px solid #1F4E79;padding-bottom:4px;margin-bottom:8px}
.to{text-align:right;line-height:1.35}h1{font-size:13.5px;text-align:center;margin:10px 0}h2{font-size:12px;color:#1F4E79;margin:10px 0 4px}
p{margin:0 0 5px;text-align:justify}table{border-collapse:collapse;width:100%;margin:2px 0 6px}th,td{border:1px solid #999;padding:3px 5px;vertical-align:top;font-size:11.5px}
th{background:#1F4E79;color:#fff;text-align:left}.g{background:#E2EFDA}ul{margin:2px 0 6px 16px;padding:0}li{margin-bottom:2px;text-align:justify}.ask{font-weight:bold;border-left:4px solid #1F4E79;padding:4px 8px;background:#f3f6fb}
.sg{margin-top:16px}"""
def tbl(head, rows, g):
    s = "<table><tr>" + "".join(f"<th>{e(h)}</th>" for h in head) + "</tr>"
    for r in rows: s += "<tr>" + "".join(f"<td{' class=g' if j in g else ''}>{e(v)}</td>" for j, v in enumerate(r)) + "</tr>"
    return s + "</table>"
h = f"""<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body>
<div class=co>{e(CO['name'])}</div><div class=rk>ИНН {CO['inn']} / КПП {CO['kpp']} · {e(CO['addr'])} · {CO['tel']} · {CO['mail']}</div>
<p>Исх. № ______ от 09.11.2026</p><div class=to>{'<br>'.join(e(x) for x in TO)}</div><h1>{e(SUBJ)}</h1>
{''.join(f'<p>{e(x)}</p>' for x in INTRO)}
{tbl(T1H, T1, {2})}{''.join(f'<p>{e(x)}</p>' for x in KEEP)}
<p class=ask>{e(ASK)}</p><p style='font-size:10px'>Приложение: спецификация АО «Тепло-Изоляционная Компания» № 15251.01 от 09.11.2026.</p>
<div class=sg>С уважением,<br>Генеральный директор ООО «ГИДРОПАС" &nbsp;&nbsp;&nbsp; ______________ &nbsp;&nbsp; /________________/<br><span style='font-size:9px'>М.П.</span></div></body></html>"""
h = h.replace("ООО «ГИДРОПАС\"", "ООО «ГИДРОПАС»")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/ppu_sogl.html"
open(hp, "w").write(h)
subprocess.run(["/opt/pw-browsers/chromium", "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer",
                f"--print-to-pdf={os.path.abspath('../PPU_UFA_SOGLASOVANIE_GOST.pdf')}", "file://" + hp], check=True, capture_output=True, timeout=180)
print("ok")
