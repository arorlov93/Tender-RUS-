"""Письмо ISP Group → АО «Алмалыкский ГМК» на новом фирменном бланке (isp_letterhead). → ../ISP_PISMO_AGMK_1011_BLANK.pdf"""
import html
from isp_letterhead import render
from isp_letter_agmk import NUM, TO, SUBJ, P1, ROWS, TERMS, ASK
e = html.escape
rows = "".join(f"<tr><td class=g>{e(a)}</td><td>{e(b).replace(chr(10), '<br>')}</td><td>{e(c)}</td></tr>" for a, b, c in ROWS)
body = f"""<div class=meta><div class=ref><b>{e(NUM.split(' от ')[0].replace('Исх. ', 'Исх. '))}</b><br>от {e(NUM.split(' от ')[1])}</div>
<div class=to><div class=l>Кому</div>{'<br>'.join(e(x) for x in TO[:3])}<br><b>{e(TO[3])}</b></div></div>
<h1>{e(SUBJ)}</h1>
<p class=greet>{e(P1[0])}</p><p>{e(P1[1])}</p>
<table><tr><th style='width:21%'>Группа</th><th>Продукция</th><th style='width:28%'>Завод-изготовитель</th></tr>{rows}</table>
<h2>Условия</h2><ul>{''.join(f'<li>{e(x[0].upper() + x[1:])}</li>' for x in TERMS)}</ul>
<p>{e(ASK)}</p>
<div class=sign><div>С уважением,<br><b style='color:#0B2545'>ISP Group LLC</b></div><div><div class=ln></div><div class=cap>подпись / М.П.</div></div>
<div class=nm><b>Александр Орлов</b>Aleksandr Orlov<br><span style='color:#6B7280;font-size:9px'>Руководитель (Member)</span></div></div>"""
render(body, "../ISP_PISMO_AGMK_1011_BLANK.pdf", "agmk2"); print("ok")
