# Где искать оптовые предложения и как проверять продавцов

## Площадки и каталоги (поиск предложений)
| Источник | Что там | Как искать |
|---|---|---|
| Parallelbroker.com | Параллельная торговля FMCG в ЕС: цены EXW (NL/PL/DE/…), MOQ (часто 33 паллеты), ежедневные обновления. Цены после регистрации | `parallelbroker.com/product/<EAN>` (Red Bull 250 мл: 9002490100056 / 9002490100070) |
| Europages, Kompass, wlw.de | Каталоги производителей и дистрибьюторов ЕС | WebSearch `site:europages.com "<товар>" wholesale distributor`; есть Apify-скрейперы Europages |
| Сайты брендов (Find a distributor) | Официальные дистрибьюторы по странам — верхняя граница цены, но надёжно | `"<бренд>" official distributor <страна>` |
| Cash & carry: Metro, Makro, Selgros, Costco | Верхняя граница оптовой цены | сайт сети страны, паллетные предложения |
| Alibaba / Tradekey / EC21 / Go4WorldBusiness / Freshdi / Globy / Exporthub | Экспортные предложения; ДЁШЕВО и МНОГО МОШЕННИКОВ | только как ориентир; поставщика — полная проверка |
| B2B-трейдеры ЕС (как VIE VITA) | FMCG-экспортёры, дилеры напитков | WebSearch `"<товар>" wholesale export price list pallets <страна>`; `"energy drinks" exporter Poland sp. z o.o.` |
| Госконтракты / тендеры (gosplan, ЕИС) | Фактические цены продаж в РФ | см. model/ (gosplan API) |

## Реестры для проверки компании
| Страна | Реестр |
|---|---|
| ЕС (все) | VIES — `scripts/check_company.py <CC> <VAT>` |
| Польша | Белый список МинФина (статус НДС + банковские счета) — в `check_company.py`; KRS: ekrs.ms.gov.pl; CEIDG/biznes.gov.pl; rejestr.io |
| Португалия | eInforma.pt, racius.com, publicacoes.mj.pt, nif.pt |
| Германия | northdata.de, handelsregister.de, unternehmensregister.de |
| Нидерланды | kvk.nl |
| Литва | registrucentras.lt, rekvizitai.vz.lt |
| Чехия / Словакия | ares.gov.cz / orsr.sk |
| Великобритания | find-and-update.company-information.service.gov.uk |
| ОАЭ | реестры фризон (DMCC, JAFZA) |
| США | sunbiz.org (FL), opencorporates.com |

## Следы и жалобы
- whois / возраст домена; совпадение домена e-mail с сайтом (gmail/outlook у «экспортёра» — флаг).
- Телефоны: совпадают ли с реестром; WhatsApp-only — флаг.
- scamadviser.com, trustpilot.com, ripoffreport.com, scamwarners.com, 419scam.org, reddit.
- Предупреждения брендов (Red Bull: policies.redbull.com … fraud_warning — «Export Scam / Non-existent Product Scam»).
