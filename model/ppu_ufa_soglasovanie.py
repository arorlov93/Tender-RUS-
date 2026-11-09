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
SUBJ = "О согласовании стандарта стальных труб для отводов 377х8 и 426х8 в ППУ-ПЭ"
INTRO = [
 "Уважаемые коллеги!",
 "ООО «ГИДРОПАС» исполняет договор на поставку отводов в ППУ-изоляции в оболочке ПЭ по ГОСТ 30732-2020 (закупка № 32616450774). "
 "Заказ подготовлен у завода-изготовителя АО «Тепло-Изоляционная Компания» (г. Серпухов).",
 "Завод сообщил, что трубы по ГОСТ 20295-85 тип 1 размеров 377х8 и 426х8 на рынке сейчас в дефиците и поставляются под заказ. "
 "Чтобы выдержать сроки поставки, просим согласовать для этих двух позиций стальные трубы по ГОСТ 8731-74 и ГОСТ 10705-80. "
 "Размеры, толщина стенки и марка стали остаются точно такими, как в техническом задании.",
]
T1H = ["№ по ТЗ", "Позиция", "Кол-во", "Размер, стенка, сталь", "По ТЗ", "Предлагаем"]
T1 = [
 ["2", "Отвод 377х8-90-1-ППУ-ПЭ, оболочка 500", "2 шт", "377 х 8,0 мм, Ст20 (без изменений)", "ГОСТ 20295-85, тип 1",
  "ГОСТ 8731-74, бесшовная горячедеформированная (основной вариант) или ГОСТ 10705-80, группа В"],
 ["3", "Отвод 426х8-90-1-ППУ-ПЭ, оболочка 560", "3 шт", "426 х 8,0 мм, Ст20 (без изменений)", "ГОСТ 20295-85, тип 1",
  "ГОСТ 10705-80, прямошовная электросварная, группа В"],
 ["1, 4", "Отводы 108х4,5 и 530х9", "3 + 1 шт", "по ТЗ", "ГОСТ 10705-80; ГОСТ 20295-85 тип 3", "без изменений, строго по ТЗ"],
]
T2H = ["Показатель", "ГОСТ 20295-85, тип 1 (по ТЗ)", "ГОСТ 10705-80, группа В", "ГОСТ 8731-74 (бесшовная)"]
T2 = [
 ["Технология", "прямошовная, сварка токами высокой частоты", "прямошовная, сварка токами высокой частоты (та же технология)", "бесшовная, горячая деформация"],
 ["Сварной шов", "есть", "есть", "нет: исключены дефекты сварного шва"],
 ["Размер и стенка", "377х8 / 426х8", "377х8 / 426х8, без изменений", "377х8, без изменений"],
 ["Марка стали", "Ст20", "Ст20 (ГОСТ 1050)", "Ст20 (ГОСТ 1050)"],
 ["Контроль химсостава и механических свойств", "да", "да (группа В)", "да (группа В)"],
 ["Спиральный шов (запрещён ТЗ)", "нет", "нет", "нет"],
 ["Уже применяется в этом договоре", "позиции 2–4", "позиция 1 (108х4,5) по ТЗ", "нет"],
]
WHY = [
 "Труба по ГОСТ 10705-80 группы В изготавливается по той же технологии (прямой шов, сварка ТВЧ), что и труба ГОСТ 20295-85 тип 1, "
 "из той же стали Ст20 с контролем химического состава и механических свойств. Этот стандарт уже предусмотрен вашим техническим заданием для позиции 1 (отвод 108х4,5).",
 "Бесшовная труба по ГОСТ 8731-74 не имеет сварного шва вообще. Для отвода, который работает на изгиб и при перепадах температур, это повышает надёжность: "
 "в изделии нет зоны, где возможны непровар, трещина или коррозия по шву.",
 "Требование ТЗ о недопустимости спиральношовных труб полностью соблюдается.",
]
KEEP = [
 "цена договора и цены за единицу не меняются;",
 "сроки и условия поставки, доставка и разгрузка в Уфе не меняются;",
 "все остальные требования ТЗ сохраняются: ГОСТ 30732-2020, отводы по ГОСТ 17375, ППУ в оболочке ПЭ, система ОДК (3 медных проводника на 377 и 426), концы 150/210 мм, изготовление 2026 года;",
 "передаём паспорта на изделия, сертификаты на трубы по ГОСТ 8731-74 / ГОСТ 10705-80 и документ завода об исправности системы ОДК на дату отгрузки.",
]
ASK = ("Просим письменно согласовать применение указанных труб для позиций 2 и 3 до 13.11.2026. "
       "Это позволит запустить изделия в производство и поставить их в срок.")

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
for i, x in enumerate(INTRO): P(x, al="j" if i else None, b=(i == 0))
P("1. Что просим согласовать", b=True, color="1F4E79")
table(T1H, T1, [1.3, 3.6, 1.4, 3.2, 2.8, 4.7], hl_col={5})
P("2. Почему предлагаемые трубы не хуже, а по ряду показателей лучше требуемых", b=True, color="1F4E79")
table(T2H, T2, [4.2, 3.9, 4.1, 4.8], hl_col={2, 3})
for x in WHY: P("•  " + x, al="j", after=2)
P("3. Что остаётся без изменений", b=True, color="1F4E79")
for x in KEEP: P("•  " + x, al="j", after=2)
P(ASK, b=True, al="j", after=10)
P("Приложение: спецификация АО «Тепло-Изоляционная Компания» № 15251.01 от 09.11.2026.", sz=10, after=14)
P("С уважением,", after=0)
P("Генеральный директор ООО «ГИДРОПАС»          ______________          /________________/", after=0)
P("М.П.", sz=9)
d.save("../PPU_UFA_SOGLASOVANIE_GOST.docx")

# ---------------- PDF (A4 портрет) ----------------
e = html.escape
css = """@page{size:A4 portrait;margin:14mm 16mm}body{font-family:'Times New Roman','DejaVu Serif',serif;font-size:11px;color:#111}
.co{font-size:20px;font-weight:bold;color:#1F4E79}.rk{font-size:8.5px;color:#555;border-bottom:2px solid #1F4E79;padding-bottom:4px;margin-bottom:8px}
.to{text-align:right;line-height:1.35}h1{font-size:13.5px;text-align:center;margin:10px 0}h2{font-size:12px;color:#1F4E79;margin:10px 0 4px}
p{margin:0 0 5px;text-align:justify}table{border-collapse:collapse;width:100%;margin:2px 0 6px}th,td{border:1px solid #999;padding:3px 5px;vertical-align:top;font-size:10px}
th{background:#1F4E79;color:#fff;text-align:left}.g{background:#E2EFDA}ul{margin:2px 0 6px 16px;padding:0}li{margin-bottom:2px;text-align:justify}.ask{font-weight:bold;border-left:4px solid #1F4E79;padding:4px 8px;background:#f3f6fb}
.sg{margin-top:16px}"""
def tbl(head, rows, g):
    s = "<table><tr>" + "".join(f"<th>{e(h)}</th>" for h in head) + "</tr>"
    for r in rows: s += "<tr>" + "".join(f"<td{' class=g' if j in g else ''}>{e(v)}</td>" for j, v in enumerate(r)) + "</tr>"
    return s + "</table>"
h = f"""<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body>
<div class=co>{e(CO['name'])}</div><div class=rk>ИНН {CO['inn']} / КПП {CO['kpp']} · {e(CO['addr'])} · {CO['tel']} · {CO['mail']}</div>
<p>Исх. № ______ от 09.11.2026</p><div class=to>{'<br>'.join(e(x) for x in TO)}</div><h1>{e(SUBJ)}</h1>
<p><b>{e(INTRO[0])}</b></p>{''.join(f'<p>{e(x)}</p>' for x in INTRO[1:])}
<h2>1. Что просим согласовать</h2>{tbl(T1H, T1, {5})}
<h2>2. Почему предлагаемые трубы не хуже, а по ряду показателей лучше требуемых</h2>{tbl(T2H, T2, {2, 3})}
<ul>{''.join(f'<li>{e(x)}</li>' for x in WHY)}</ul>
<h2>3. Что остаётся без изменений</h2><ul>{''.join(f'<li>{e(x)}</li>' for x in KEEP)}</ul>
<p class=ask>{e(ASK)}</p><p style='font-size:10px'>Приложение: спецификация АО «Тепло-Изоляционная Компания» № 15251.01 от 09.11.2026.</p>
<div class=sg>С уважением,<br>Генеральный директор ООО «ГИДРОПАС" &nbsp;&nbsp;&nbsp; ______________ &nbsp;&nbsp; /________________/<br><span style='font-size:9px'>М.П.</span></div></body></html>"""
h = h.replace("ООО «ГИДРОПАС\"", "ООО «ГИДРОПАС»")
hp = "/tmp/claude-0/-home-user-Tender-RUS-/e981d7a6-b89f-5df6-824c-c1907918e938/scratchpad/ppu_sogl.html"
open(hp, "w").write(h)
subprocess.run(["/opt/pw-browsers/chromium", "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer",
                f"--print-to-pdf={os.path.abspath('../PPU_UFA_SOGLASOVANIE_GOST.pdf')}", "file://" + hp], check=True, capture_output=True, timeout=180)
print("ok")
