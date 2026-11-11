"""Официальный запрос ISP Group → СП ООО «Uz-Kor Gas Chemical» на тендерную документацию T-47/2026. → ../ISP_ZAPROS_UZKOR_T47_1111.pdf"""
import html, os, sys
SIG = sys.argv[1] if len(sys.argv) > 1 else ''
from isp_letterhead import render
e = html.escape
NUM, DATE = "ISP-1111-02", "11.11.2026"
TO = ["Отдел закупок", "СП ООО «Uz-Kor Gas Chemical»", "Республика Узбекистан", "Касимову А."]
SUBJ = "Запрос тендерной документации по тендеру № T-47/2026"
P = ["Уважаемый господин Касимов!",
     "ISP Group LLC (США) просит предоставить тендерную инструкцию и техническое задание по тендеру № T-47/2026 на поставку химических реагентов для цеха ОЗХ.",
     "Наша компания поставляет реагенты для водоподготовки и нефтегазохимической отрасли предприятиям Центральной Азии и работает напрямую с заводами-изготовителями в Китае. Мы намерены подать предложение в установленный срок до 02.12.2026."]
ROWS = [("Наименование", "ISP Group LLC"),
        ("Регистрация", "Wyoming LLC, зарегистрирована как иностранная компания во Флориде, Doc. No. M26000003981, EIN 38-4380828"),
        ("Адрес", "16395 Biscayne Blvd, North Miami Beach, FL 33160, USA"),
        ("Руководитель", "Александр Орлов, Managing Member"),
        ("Контакты", "info@ispgroupgc.com, +1 929 830 0038, ispgroupgc.com"),
        ("Продукция", "Ингибиторы солеотложений и коррозии (ОЭДФ, НТФ, ФБТК, DTPMPA), антискаланты, реагенты для оборотной воды, коагулянты (PAC), ионообменные смолы")]
ASK = "Просим направить тендерную документацию на адрес info@ispgroupgc.com. Если для участия требуется регистрация или дополнительные документы, просим сообщить порядок."
rows = "".join(f"<tr><td class=g>{e(a)}</td><td>{e(b)}</td></tr>" for a, b in ROWS)
SIGIMG = f"<img src='file://{SIG}' style='position:absolute;left:18%;bottom:-6px;width:1.05in'>" if SIG else ""
body = f"""<div class=meta><div class=ref><b>Исх. № {NUM}</b><br>от {DATE}</div>
<div class=to><div class=l>Кому</div>{'<br>'.join(e(x) for x in TO[:3])}<br><b>{e(TO[3])}</b></div></div>
<h1>{e(SUBJ)}</h1>
<p class=greet>{e(P[0])}</p><p>{e(P[1])}</p><p>{e(P[2])}</p>
<h2>Сведения об участнике</h2>
<table><tr><th style='width:24%'>Параметр</th><th>Значение</th></tr>{rows}</table>
<p>{e(ASK)}</p>
<div class=sign><div>С уважением,<br><b style='color:#0B2545'>ISP Group LLC</b></div><div><div class=ln style='position:relative'>{SIGIMG}</div><div class=cap>подпись / М.П.</div></div>
<div class=nm><b>Александр Орлов</b>Aleksandr Orlov<br><span style='color:#6B7280;font-size:9px'>Руководитель (Member)</span></div></div>"""
render(body, "../ISP_ZAPROS_UZKOR_T47_1111" + ("_SIGNED" if SIG else "") + ".pdf", "uzkor"); print("ok")
