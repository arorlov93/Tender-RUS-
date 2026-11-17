# Скрипт для Claude in Chrome: проверка профиля SAM.gov и поиск закупок

Как пользоваться: войдите на https://sam.gov (Login.gov + код), откройте Claude in Chrome на этой вкладке и вставьте блок ниже целиком. Claude только смотрит и собирает данные — ничего не меняет и не отправляет.

---

```
Ты работаешь в моём браузере на sam.gov, я уже залогинен. Компания: ISP Group LLC (Флорида).
ПРАВИЛА: ничего не редактируй, не сохраняй, не продлевай, не подавай и не нажимай Submit/Save/Update/Renew/Certify. Только читай. Если сессия истекла или нужен код входа — остановись и скажи мне.

ЧАСТЬ 1. Проверка регистрации (Workspace → Entities → ISP Group LLC → View/Entity Registration):
1. Выпиши: UEI, CAGE, статус регистрации (Active/Inactive/Submitted/Expired), дату активации и дату истечения (Expiration Date), Purpose of Registration (должно быть "All Awards", а не только "Federal Assistance Awards").
2. Раздел Goods & Services: выпиши все NAICS-коды, какой из них Primary, и отметку Small Business (Yes/No) по каждому. Выпиши PSC-коды, если указаны.
3. Сравни со списком нужных кодов и отметь ЕСТЬ / НЕТ / малый бизнес:
   424690 (Other Chemical and Allied Products Merchant Wholesalers) — должен быть Primary
   325180 (Other Basic Inorganic Chemical Manufacturing)
   325998 (All Other Miscellaneous Chemical Product Manufacturing)
   325199 (All Other Basic Organic Chemical Manufacturing)
   424610 (Plastics Materials and Basic Forms and Shapes Merchant Wholesalers)
   424720 / 423830 / 423840 — если есть, просто выпиши
4. Reps & Certs (FAR 52.204-8 / 52.212-3): выпиши ответы по пунктам — small business, ownership/control (Owned or controlled by a foreign... / immediate & highest-level owner), 52.204-24/26 (Section 889), 52.209-11/-5 (налоговая задолженность/судимость), 52.225-2/-4/-6 (Buy American / TAA — указанные страны происхождения товаров), 52.204-17/-20 (материнская компания). Ничего не меняй — только выпиши.
5. Points of Contact: заполнены ли E-Business POC и Government Business POC (только да/нет, без персональных данных в ответе).
6. Нет ли красных/жёлтых уведомлений, Exclusions, задач "Action required".

ЧАСТЬ 2. Поиск закупок (Search → Contract Opportunities), фильтр Status: Active, Notice type: Solicitation + Combined Synopsis/Solicitation + Sources Sought + Presolicitation:
Сделай отдельные поиски и по каждому выпиши все результаты с ответом позже чем через 5 дней от сегодня:
 A. NAICS 325180, 325998, 424690, 325199, 325188
 B. PSC 6810 (Chemicals), 6850 (Miscellaneous Chemical Specialties), 6840 (Pest control / disinfectants), 4610 (Water purification equipment) — только если речь о химии/расходниках
 C. Ключевые слова (каждое отдельно): "potassium permanganate", "sodium permanganate", "activated carbon", "granular activated carbon", "powdered activated carbon", "water treatment chemicals", "coagulant", "polymer" + "water", "polyaluminum chloride", "ferric chloride", "corrosion inhibitor", "demulsifier", "ion exchange resin", "antiscalant"
Для каждой закупки: Notice ID, название, ведомство (Department/Sub-tier/Office), место поставки (штат), тип уведомления, set-aside (Total Small Business / 8(a) / SDVOSB / нет), NAICS, срок ответа (дата+часовой пояс), объём/количество если видно в описании, требования NSF/ANSI 60/61, AWWA, Buy American/TAA, ссылка.
Убери дубли. Отметь красным закупки с set-aside 8(a), SDVOSB, WOSB, HUBZone — нам недоступны. Отметь закупки, где товар — бочки/кубы/мешки, а не налив в танкерах.

ЧАСТЬ 3. Итог: дай мне
 1) таблицу регистрации (поле — значение — ок/проблема),
 2) список недостающих NAICS,
 3) таблицу закупок, отсортированную по сроку подачи,
 4) экспорт таблицы закупок в CSV (скопируй как текст в чат).
```

---

После этого пришлите мне вывод (или CSV) — посчитаю цены и маржу по каждой закупке и подготовлю пакеты.

Если в части 1 окажется, что NAICS не хватает: добавление кода — это правка регистрации (Update Registration), она снова уходит на проверку. Делайте её сами, вне скрипта; на время проверки регистрация остаётся активной.
