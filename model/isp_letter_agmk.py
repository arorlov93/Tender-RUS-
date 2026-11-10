"""Официальное письмо ISP Group LLC → АО «Алмалыкский ГМК» (по запросу Халдаровой Г. Х., 10.11.2026): предложение о поставке
ионообменных смол и фосфонатов, на русском, PDF + DOCX, место для подписи руководителя. Банковских реквизитов нет.
→ ../ISP_PISMO_AGMK_1011.pdf / .docx"""
import html, os, subprocess
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

HEAD = "ISP Group LLC"
REQ = ("16395 Biscayne Blvd, North Miami Beach, FL 33160, USA  ·  EIN 38-4380828  ·  Florida Doc. No. M26000003981  ·  "
       "+1 929 830 0038  ·  info@ispgroupgc.com  ·  ispgroupgc.com")
NUM = "Исх. № ISP-1011-01 от 10.11.2026"
TO = ["АО «Алмалыкский горно-металлургический комбинат»", "Службе закупок", "Халдаровой Г. Х.", "info@agmk.uz"]
SUBJ = "Предложение о поставке ионообменных смол и реагентов водоподготовки"
P1 = ("Уважаемая Гулшода Хасан кизи!",
      "Компания ISP Group LLC (США) предлагает АО «Алмалыкский ГМК» сотрудничество в поставке ионообменных смол и реагентов "
      "для водоподготовки. Мы работаем напрямую с двумя крупными заводами-изготовителями в Китае.")
ROWS = [
 ("Ионообменные смолы", "Сильнокислотный катионит гелевого типа (St-DVB), H- и Na-форма (аналог КУ-2-8)\n"
  "Сильноосновный анионит гелевого и макропористого типа (St-DVB), Cl-форма (аналог АВ-17-8)\nСлабоосновный макропористый анионит",
  "Завод основан в 1970 г., мощность более 50 000 м³ смол в год; ISO 9001, ISO 14001, ISO 45001"),
 ("Фосфонаты и ингибиторы солеотложений", "НТФ (ATMP) 50 % и 95 % сухая; ОЭДФ (HEDP) 60 % и 90 % сухая; PBTCA 50 %; DTPMPA 50 %;\n"
  "антискалант для установок обратного осмоса", "Мощность завода 280 000 т в год; ISO 9001, ISO 14001"),
]
TERMS = [
 "на каждую марку предоставляем паспорт качества (COA), техническое описание (TDS) и паспорт безопасности (SDS);",
 "образцы для испытаний направим бесплатно;",
 "поставка морем или ж/д транспортом, полными или сборными контейнерами, на условиях FOB Китай или с доставкой до вашего склада;",
 "упаковка: смолы в мешках по 25 л; реагенты в мешках 25 кг, бочках 250 кг или еврокубах 1000–1250 кг;",
 "срок изготовления 20–30 дней с момента подтверждения заказа.",
]
ASK = ("Просим сообщить, какие из указанных продуктов закупает комбинат, ориентировочный годовой объём и порядок участия в закупках "
       "(электронная площадка, регистрация поставщика). По вашему запросу подготовим коммерческое предложение с ценами и сроками поставки.")
SIGN = ("Руководитель (Member) ISP Group LLC", "А. Орлов / Aleksandr Orlov")

# DOCX
d = Document()
for s in d.sections: s.left_margin = s.right_margin = Cm(2); s.top_margin = s.bottom_margin = Cm(1.5)
st = d.styles["Normal"]; st.font.name = "Times New Roman"; st.font.size = Pt(11)
def P(t, b=False, al=None, sz=None, c=None, after=4):
    p = d.add_paragraph(); r = p.add_run(t); r.bold = b
    if sz: r.font.size = Pt(sz)
    if c: r.font.color.rgb = RGBColor.from_string(c)
    if al: p.alignment = {"c": WD_ALIGN_PARAGRAPH.CENTER, "r": WD_ALIGN_PARAGRAPH.RIGHT, "j": WD_ALIGN_PARAGRAPH.JUSTIFY}[al]
    p.paragraph_format.space_after = Pt(after)
P(HEAD, b=True, sz=18, c="1F3864", after=0); P(REQ, sz=8, c="555555", after=10)
P(NUM, sz=10)
for x in TO: P(x, al="r", after=0)
P("", after=4); P(SUBJ, b=True, al="c", sz=12.5, after=8)
P(P1[0], b=True); P(P1[1], al="j")
t = d.add_table(rows=1, cols=3); t.style = "Table Grid"
for j, h in enumerate(["Группа", "Продукция", "Завод-изготовитель"]): t.rows[0].cells[j].text = h
for r in ROWS:
    cs = t.add_row().cells
    for j, v in enumerate(r): cs[j].text = v
for row in t.rows:
    for j, w in enumerate([3.5, 9.5, 4.5]): row.cells[j].width = Cm(w)
P("", after=2); P("Условия:", b=True, after=2)
for x in TERMS: P("•  " + x, al="j", after=1)
P(ASK, al="j", after=14)
P("С уважением,", after=0); P(f"{SIGN[0]}          ____________________          {SIGN[1]}", after=0)
d.save("../ISP_PISMO_AGMK_1011.docx")

# PDF
e = html.escape
css = """@page{size:A4 portrait;margin:15mm 17mm}body{font-family:'Times New Roman','DejaVu Serif',serif;font-size:12.5px;color:#111}
.h{font-size:24px;font-weight:bold;color:#1F3864}.rq{font-size:8.5px;color:#555;border-bottom:2px solid #1F3864;padding-bottom:5px;margin-bottom:12px}
.to{text-align:right;line-height:1.4}h1{font-size:14px;text-align:center;margin:14px 0}p{margin:0 0 7px;text-align:justify}
table{border-collapse:collapse;width:100%;margin:4px 0 8px}th,td{border:1px solid #999;padding:4px 6px;vertical-align:top;font-size:11.5px}th{background:#1F3864;color:#fff;text-align:left}
ul{margin:2px 0 8px 18px;padding:0}li{margin-bottom:2px}.sg{margin-top:26px;display:flex;justify-content:space-between;align-items:flex-end}.ln{border-bottom:1px solid #000;width:170px;height:28px}"""
tb = "".join(f"<tr><td><b>{e(a)}</b></td><td>{e(b).replace(chr(10), '<br>')}</td><td>{e(c)}</td></tr>" for a, b, c in ROWS)
h = f"""<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body>
<div class=h>{HEAD}</div><div class=rq>{e(REQ)}</div><p>{NUM}</p><div class=to>{'<br>'.join(e(x) for x in TO)}</div><h1>{e(SUBJ)}</h1>
<p><b>{e(P1[0])}</b></p><p>{e(P1[1])}</p>
<table><tr><th style='width:20%'>Группа</th><th>Продукция</th><th style='width:27%'>Завод-изготовитель</th></tr>{tb}</table>
<p><b>Условия:</b></p><ul>{''.join(f'<li>{e(x)}</li>' for x in TERMS)}</ul><p>{e(ASK)}</p>
<div class=sg><div>С уважением,<br>{e(SIGN[0])}</div><div class=ln></div><div>{e(SIGN[1])}</div></div></body></html>"""
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/agmk.html"
open(hp, "w").write(h)
subprocess.run(["/opt/pw-browsers/chromium", "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer",
                f"--print-to-pdf={os.path.abspath('../ISP_PISMO_AGMK_1011.pdf')}", "file://" + hp], check=True, capture_output=True, timeout=180)

# ---- на фирменном бланке ISP Group (Letter): шапка и подвал из бланка, тело письма поверх ----
import sys, re, pypdf
BLANK = sys.argv[1] if len(sys.argv) > 1 else None
if BLANK:
    css2 = css.replace("@page{size:A4 portrait;margin:15mm 17mm}", "@page{size:Letter portrait;margin:40mm 22mm 17mm 16mm}").replace("font-size:12.5px", "font-size:12px")
    h2 = re.sub(r"<div class=h>.*?</div><div class=rq>.*?</div>", "", h.replace(css, css2), flags=re.S)
    hp2 = hp.replace(".html", "_body.html"); open(hp2, "w").write(h2)
    body = hp2.replace(".html", ".pdf")
    subprocess.run(["/opt/pw-browsers/chromium", "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer",
                    f"--print-to-pdf={body}", "file://" + hp2], check=True, capture_output=True, timeout=180)
    r = pypdf.PdfReader(BLANK); pg = r.pages[0]
    raw = pg.get_contents().get_data().decode("latin-1")
    head = raw[:raw.index("Q\n") + 2]                      # первый блок q…Q: шапка + линия + подвал; поля Date/To/Subject и текст бланка убираем
    from pypdf.generic import DecodedStreamObject, NameObject
    st_ = DecodedStreamObject(); st_.set_data(head.encode("latin-1")); pg[NameObject("/Contents")] = r._writer_add_obj(st_) if hasattr(r, "_writer_add_obj") else st_
    w = pypdf.PdfWriter(); w.add_page(pg)
    bp = pypdf.PdfReader(body).pages[0]; w.pages[0].merge_page(bp)
    w.write("../ISP_PISMO_AGMK_1011_BLANK.pdf")
print("ok")
