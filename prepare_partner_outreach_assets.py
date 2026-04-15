from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill


WORKBOOK_PATH = Path("/workspace/Сколково контакты.xlsx")
TARGET_SHEETS = ["Екатеринбург", "Казань", "Нижний Новгород", "Краснодарский край"]
OUTPUT_SHEET = "Влад - объединено"
EXCLUDED_RESPONSIBLES = ("аня", "анна", "лена", "елена", "юля")
ADDED_FONT = Font(color="666666")
HEADER_FILL = PatternFill(fill_type="solid", fgColor="DCE6F1")


@dataclass(frozen=True)
class AddedInfo:
    email_or_form: str = ""
    phone: str = ""
    social: str = ""
    official_page: str = ""
    confidence: str = ""
    note: str = ""
    sources: str = ""


ADDED_CONTACTS: dict[tuple[str, int], AddedInfo] = {
    ("Екатеринбург", 2): AddedInfo(
        email_or_form="asp@altekproekt.ru",
        phone="+7 (343) 283-07-37; +7 (343) 283-07-30",
        official_page="ООО «Альтек Строй Проект»: altek-stroi-proekt.eb24.ru",
        confidence="средняя-высокая",
        note="Личный TG уже есть в базе. Добавлен корпоративный путь через компанию.",
        sources="https://altek-stroi-proekt.eb24.ru/ ; https://companies.rbc.ru/id/1106674009122-ooo-altek-stroj-proekt/",
    ),
    ("Екатеринбург", 3): AddedInfo(
        email_or_form="expertural@expertural.com",
        phone="+7 (343) 345-03-42",
        official_page="Контакты «Эксперт-Урал»: expert-ural.com/kontakti/ ; страница автора: expert-ural.com/authors/kovalenko-svetlana.html",
        confidence="высокая",
        note="Для первого касания безопаснее идти через редакционный e-mail и ссылку на страницу Светланы.",
        sources="https://expert-ural.com/kontakti/ ; https://expert-ural.com/authors/kovalenko-svetlana.html",
    ),
    ("Екатеринбург", 4): AddedInfo(
        official_page="Официальный сайт бренда: yarmysheva.ru",
        confidence="средняя",
        note="Надежный публичный прямой e-mail не подтвержден; лучше заходить через сайт бренда.",
        sources="https://yarmysheva.ru/",
    ),
    ("Екатеринбург", 5): AddedInfo(
        confidence="низкая",
        note="Без ИНН/точного юрлица надежный публичный контакт не найден; высокий риск ошибки по одноименным компаниям.",
    ),
    ("Екатеринбург", 6): AddedInfo(
        email_or_form="aviharev-vc@yandex.ru",
        phone="+7 (922) 130-44-33; +7 (912) 207-19-54",
        social="VK: vk.com/volunteersekb",
        official_page="Контакты Волонтерского центра",
        confidence="высокая",
        note="Подходит для аккуратного выхода через центр/приемную.",
        sources="https://xn----8sbafbbrdz0awkb4aez0d3f.xn--p1ai/kontakty/ ; https://vk.com/volunteersekb",
    ),
    ("Екатеринбург", 7): AddedInfo(
        phone="+7 (343) 226-02-72",
        official_page="Карточка ООО «КВАРТА»",
        confidence="средняя",
        note="Нужно перепроверить ИНН/юрлицо: в открытых базах встречаются одноименные компании.",
        sources="https://companium.ru/id/1136678003252-kvarta",
    ),
    ("Екатеринбург", 8): AddedInfo(
        email_or_form="fr@lifemart.ru; pr@lifemart.ru",
        phone="+7 912 044-89-64; +7 (800) 700-56-11",
        official_page="Lifemart / франшиза; общий канал сети «Сушкоф»",
        confidence="высокая",
        note="Для партнерства лучше писать через Lifemart fr/pr, а не в лоб на общий контакт сети.",
        sources="https://fr.lifemart.ru/ ; https://www.orgpage.ru/ekaterinburg/sushkof-5973488.html",
    ),
    ("Екатеринбург", 9): AddedInfo(
        email_or_form="Dmitry.Larin@okkam.ru",
        official_page="Профиль команды Okkam и страница контактов",
        confidence="средняя-высокая",
        note="E-mail найден в публичной индексации; перед массовой рассылкой стоит визуально перепроверить страницу.",
        sources="https://okkam.group/team/dmitrij_larin ; https://okkam.group/contacts",
    ),
    ("Екатеринбург", 10): AddedInfo(
        email_or_form="post@bergauf.ru",
        phone="+7 (343) 278-52-94; +7 (343) 310-79-35",
        official_page="Bergauf contacts; региональные контакты Isover/ Saint-Gobain",
        confidence="средняя",
        note="Личного публичного контакта Ирины не найдено; добавлены брендовые каналы компаний.",
        sources="https://bergauf.ru/contacts/ ; https://www.orgpage.ru/ekaterinburg/isover-proizvodstvennaya-1864503.html",
    ),
    ("Екатеринбург", 11): AddedInfo(
        official_page="Карточка ООО ИК «Энергософт» с ИНН и адресом",
        confidence="высокая",
        note="В базе уже есть прямой телефон, TG и сайт; дополнительно подтверждена связка с юрлицом.",
        sources="https://companies.rbc.ru/id/1126671019661-ooo-inzhiniringovaya-kompaniya-energosoft/",
    ),
    ("Екатеринбург", 12): AddedInfo(
        confidence="низкая",
        note="В базе уже есть прямой TG. Дополнительный надежный публичный канал без риска ошибки не найден.",
    ),
    ("Екатеринбург", 13): AddedInfo(
        confidence="низкая",
        note="В базе уже есть прямой TG. Для точного web-enrichment нужен ИНН одной из компаний.",
    ),
    ("Екатеринбург", 14): AddedInfo(
        email_or_form="info@maritol.ru",
        phone="+7 (343) 268-30-61",
        official_page="maritol.ru/about/",
        confidence="высокая",
        note="Добавлен надежный корпоративный контакт компании.",
        sources="https://maritol.ru/about/",
    ),
    ("Екатеринбург", 15): AddedInfo(
        phone="+7 (343) 363-53-25; +7 (922) 020-53-94",
        official_page="autoprocat.ru и карточка 2GIS",
        confidence="средняя",
        note="В базе уже есть прямой TG; e-mail с сайта не подтвердился, поэтому оставлен телефон/сайт.",
        sources="https://autoprocat.ru/ ; https://2gis.ru/firm/70000001038545258",
    ),
    ("Екатеринбург", 16): AddedInfo(
        social="@MyRyadom2020bot",
        official_page="miryadom2020.tilda.ws ; xn--2020-p4d1cbpq3iyb.xn--p1ai",
        confidence="средняя-высокая",
        note="В базе уже есть прямой TG; дополнительно подтверждены официальные страницы проекта.",
        sources="https://miryadom2020.tilda.ws/ ; https://xn--2020-p4d1cbpq3iyb.xn--p1ai/",
    ),
    ("Екатеринбург", 33): AddedInfo(
        phone="+7 (343) 286-11-63",
        official_page="Уралэкспоцентр: uralex.ru/index.php/contacts",
        confidence="низкая",
        note="Только условный корпоративный путь; использовать после ручной проверки, что это именно нужный Александр Баранов.",
        sources="https://www.uralex.ru/index.php/contacts",
    ),
    ("Казань", 2): AddedInfo(
        official_page="Карточка ООО «РЕСУРС» с ИНН/ОГРН и адресом",
        confidence="средняя",
        note="Подтверждена связка с юрлицом, но открытый телефон/e-mail не найден.",
        sources="https://companies.rbc.ru/id/1211600091625-obschestvo-s-ogranichennoj-otvetstvennostyu-resurs/",
    ),
    ("Казань", 6): AddedInfo(
        phone="8 (800) 700-03-88; +7 (843) 250-15-24",
        official_page="marinaizmaylova.ru ; адрес: Казань, ул. Дорожная, 12а",
        confidence="средняя",
        note="Личный e-mail с лендинга не подтвердился; добавлены официальный сайт и корпоративные телефоны из открытых справочников.",
        sources="https://marinaizmaylova.ru/ ; https://mestam.info/ru/kazan/mesto/979036-standartprodmash-dorojnaya-12a",
    ),
    ("Казань", 7): AddedInfo(
        email_or_form="eco@tatar.ru; интернет-приемная",
        phone="+7 (843) 267-68-01; +7 (843) 267-68-02",
        official_page="Контакты Минэкологии РТ",
        confidence="высокая",
        note="Для такого контакта разумнее идти через приемную/официальный канал.",
        sources="https://eco.tatarstan.ru/kontaktnaya-informatsiya.htm ; https://eco.tatarstan.ru/rus/priem.htm",
    ),
    ("Казань", 8): AddedInfo(
        email_or_form="info@romangroup.ru",
        phone="+7 (843) 203-93-87",
        social="VK федерации: vk.com/fctat",
        official_page="ROMAN GROUP contacts; Федерация скалолазания РТ",
        confidence="высокая",
        note="Есть два рабочих пути: через бизнес и через спортивную федерацию.",
        sources="https://romangroup.ru/contacts/ ; https://www.romangroup.ru/company/staff/yakovlev-stanislav-igorevich/ ; https://vk.com/fctat",
    ),
    ("Казань", 10): AddedInfo(
        confidence="низкая",
        note="Без точного юрлица/ИНН не удалось надежно сопоставить контакт.",
    ),
    ("Казань", 11): AddedInfo(
        email_or_form="info@tssenergy.ru",
        phone="+7 (843) 590-36-30",
        official_page="tssenergy.ru",
        confidence="высокая",
        note="Добавлен корпоративный контакт компании; личная привязка требует ручной проверки.",
        sources="http://tssenergy.ru/",
    ),
    ("Казань", 14): AddedInfo(
        email_or_form="tat-chess@mail.ru",
        phone="+7 (843) 236-58-26",
        official_page="Федерация шахмат РТ",
        confidence="средняя",
        note="Добавлен официальный контакт федерации как безопасный вход.",
        sources="https://tat-chess.ru/ ; https://federatsiya-shakhmat-respubliki.orgs.biz/",
    ),
    ("Казань", 16): AddedInfo(
        official_page="Карточка ликвидированного ИП",
        confidence="средняя",
        note="Актуальный публичный контакт не найден; в открытых источниках подтверждается только ликвидация ИП.",
        sources="https://datanewton.ru/contragents/306168515300031",
    ),
    ("Казань", 17): AddedInfo(
        email_or_form="Evadesign@internet.ru",
        official_page="evadesignhome.ru/kontakty",
        confidence="средняя",
        note="В открытых данных найден бренд EVADESIGN; фамилию/юрлицо нужно перепроверить перед отправкой.",
        sources="https://companies.rbc.ru/id/1161690168771-ooo-evo-dizajn/ ; https://evadesignhome.ru/kontakty",
    ),
    ("Казань", 20): AddedInfo(
        email_or_form="tpprt@tpprt.ru",
        phone="+7 (843) 264-62-07",
        official_page="ТПП Татарстана",
        confidence="низкая",
        note="Это скорее ориентир по деловому сообществу, чем подтвержденный прямой контакт Елены Кузнецовой.",
        sources="https://tatarstan.tpprf.ru/ru/contacts/",
    ),
    ("Казань", 21): AddedInfo(
        email_or_form="rgajazov@gmail.com",
        official_page="Карточка ИП/публичный профиль",
        confidence="средняя",
        note="Название клиники в открытых источниках не подтвердилось; использовать аккуратно и лучше сначала в соцсетях.",
        sources="https://companies.rbc.ru/persons/ogrnip/319169000167442-turisticheskaya-kompaniya-7daytravel/ ; https://companium.ru/people/inn/165714096844-gayazov-rustem-rinatovich",
    ),
    ("Казань", 23): AddedInfo(
        email_or_form="ildar.mn82@gmail.com",
        official_page="Карточки ТРИТОН ИНВЕСТ / ПРОГРЕСС ОЙЛ",
        confidence="средняя",
        note="Публичный e-mail найден в открытом досье; перед email-рассылкой лучше перепроверить по дополнительным источникам.",
        sources="https://companies.rbc.ru/id/1211600016319-obschestvo-s-ogranichennoj-otvetstvennostyu-triton-invest/ ; https://companies.rbc.ru/id/1111690086837-ooo-progress-ojl/ ; https://reputation.ru/ogrn/1211600016319",
    ),
    ("Казань", 25): AddedInfo(
        email_or_form="MyCornerMarketing@yandex.ru",
        phone="+7 (800) 600-17-25",
        social="VK: vk.com/mycorner_by_unistroy",
        official_page="mycorner.ru",
        confidence="высокая",
        note="Хороший рабочий путь для партнерского предложения через бренд/маркетинг.",
        sources="https://mycorner.ru/ ; https://realnoevremya.ru/persons/2174-kornilov-anton-yurevich ; https://kazan.spravochnik-rf.ru/stroitelnye-kompanii/949900.html",
    ),
    ("Казань", 29): AddedInfo(
        email_or_form="ABIReception@abdev.ru",
        phone="+7 (843) 205-49-90",
        official_page="akbars-eng.ru",
        confidence="средняя",
        note="В открытых данных есть действующий корпоративный путь через «АК БАРС Инжиниринг», но с вашей пометкой «ликвидирована» нужно сверить, тот ли это человек.",
        sources="https://akbars-eng.ru/ ; https://reputation.ru/inn/166102619176",
    ),
    ("Казань", 33): AddedInfo(
        confidence="средняя",
        note="В базе уже есть TG; дополнительный подтвержденный деловой канал не найден.",
    ),
    ("Казань", 34): AddedInfo(
        official_page="Профиль executive.ru",
        confidence="низкая",
        note="Совпадение с казанским контактом не подтверждено; использовать только после ручной проверки.",
        sources="https://www.e-xecutive.ru/users/1798220-uliya-matunina",
    ),
    ("Казань", 35): AddedInfo(
        official_page="Карточка ИП с совпадающим ФИО",
        confidence="низкая",
        note="Однозначная привязка к нужному человеку не подтверждена.",
        sources="https://companies.rbc.ru/persons/ogrnip/315169000020962-husnutdinov-ajdar-shajhelislamovich/",
    ),
    ("Казань", 37): AddedInfo(
        official_page="imgevent.ru/o-nas",
        confidence="высокая",
        note="В базе уже есть сайт агентства; отдельный публичный e-mail конкретно Виталия быстро не нашелся.",
        sources="https://imgevent.ru/o-nas",
    ),
    ("Казань", 38): AddedInfo(
        official_page="imgevent.ru/o-nas",
        confidence="средняя",
        note="Нужна ручная проверка фамилии/идентификации, но как вход можно использовать общий сайт агентства.",
        sources="https://imgevent.ru/o-nas",
    ),
    ("Нижний Новгород", 2): AddedInfo(
        email_or_form="vita@kis.ru",
        phone="+7 (831) 412-32-17; +7 (910) 393-47-17",
        official_page="vita-print.com",
        confidence="высокая",
        note="Надежный корпоративный путь по компании «Вита-Принт».",
        sources="https://reputation.ru/ogrn/1025203728252 ; https://novgorod.spravker.ru/proizvodstvennyie-predpriyatiya/vita-print.htm ; https://vita-print.com/",
    ),
    ("Нижний Новгород", 6): AddedInfo(
        email_or_form="zakaz@globaltest.ru",
        phone="+7 (83130) 6-77-77 доб. 153",
        official_page="globaltest.ru/kontakty/",
        confidence="высокая",
        note="В базе уже есть мобильный номер; добавлен надежный корпоративный канал для Никиты Козлова.",
        sources="https://www.elec.ru/catalog/globaltest/contacts/ ; https://globaltest.ru/kontakty/",
    ),
    ("Нижний Новгород", 7): AddedInfo(
        phone="8 (800) 500-49-21; +7 (986) 751-15-76",
        official_page="johntruck.ru/contacts/",
        confidence="высокая",
        note="Прямой e-mail не найден; безопаснее идти через контакты/форму John Truck.",
        sources="https://johntruck.ru/contacts/ ; https://www.tbank.ru/business/contractor/legal/1125263004052/",
    ),
    ("Краснодарский край", 2): AddedInfo(
        email_or_form="Форма партнерства на gastreet.com/partners",
        phone="8 (800) 700-93-20",
        official_page="Gastreet partnership page",
        confidence="высокая",
        note="Для такого контакта лучше идти через официальный партнерский вход и уже внутри просить интро на Евгению.",
        sources="https://gastreet.com/partners ; https://2021.gastreet.com/contacts",
    ),
    ("Краснодарский край", 3): AddedInfo(
        social="TG: @Mantera_career",
        official_page="mantera.ru/contacts/ ; блок пресс-службы",
        confidence="средняя",
        note="Личный e-mail Вадима в открытом виде не найден; рабочий путь — пресс-служба/корпоративные контакты группы.",
        sources="https://mantera.ru/contacts/ ; https://mantera.ru/company/ ; https://t.me/Mantera_career",
    ),
    ("Краснодарский край", 6): AddedInfo(
        phone="+7 (961) 587-77-28",
        official_page="Карточка ООО «Альфа-Сервис»",
        confidence="средняя",
        note="Нужно перепроверить ИНН: в открытых базах есть несколько одноименных ООО.",
        sources="https://companies.rbc.ru/id/1122367000733-ooo-alfa-servis/ ; https://sochi.ruspravochnik.com/company/alfa-servis-25/alfa-servis-rossiya-krasnodarskiy-kray-sochi-ulica-yana-fabriciusa-228",
    ),
    ("Краснодарский край", 7): AddedInfo(
        email_or_form="infocenter@kpresort.ru; order@kpresort.ru; sales@kpresort.ru",
        phone="+7 (800) 550-20-20; +7 (862) 245-50-50",
        social="TG: @kp_resort",
        official_page="Контакты курорта «Красная Поляна»",
        confidence="высокая",
        note="Для партнера логичнее писать на infocenter/sales, чем искать персональный адрес.",
        sources="https://ww3.krasnayapolyanaresort.ru/kurort/about/contacts ; https://krasnayapolyanaresort.ru/contact ; https://mantera.ru/company/",
    ),
    ("Краснодарский край", 8): AddedInfo(
        phone="+7 (862) 225-55-35",
        official_page="Карточка ООО «И-ТУР»",
        confidence="средняя-высокая",
        note="Официальный сайт/e-mail не подтвердились, но телефон и связка с юрлицом выглядят рабочими.",
        sources="https://companies.rbc.ru/id/1132366001580-ooo-i-tur/ ; https://sochi.kitabi.ru/firms/i-tur-nesebrskaya-ulica-6",
    ),
    ("Краснодарский край", 9): AddedInfo(
        official_page="alias-group.ru/contacts",
        confidence="высокая",
        note="На странице контактов есть e-mail и TG, но в автоматической выгрузке текст адресов не раскрылся; лучше открыть страницу вручную перед отправкой.",
        sources="https://alias-group.ru/contacts ; https://alias-group.ru/ ; https://companium.ru/id/1212300024640-ehlias-grupp",
    ),
    ("Краснодарский край", 11): AddedInfo(
        confidence="низкая",
        note="Без компании/ИНН надежный публичный контакт не найден.",
    ),
    ("Краснодарский край", 16): AddedInfo(
        email_or_form="bsfc@bsfc.com",
        phone="8 (800) 700-1-800",
        official_page="bsfc.com/about/contacts/",
        confidence="высокая",
        note="Хороший официальный вход через агентство/группу.",
        sources="https://bsfc.com/about/contacts/",
    ),
    ("Краснодарский край", 17): AddedInfo(
        official_page="Карточка ООО «Ванвин» с адресом",
        confidence="высокая",
        note="Компания подтверждена, но публичный e-mail/телефон не найден.",
        sources="https://companies.rbc.ru/id/1202300013145-obschestvo-s-ogranichennoj-otvetstvennostyu-vanvin/",
    ),
    ("Краснодарский край", 21): AddedInfo(
        email_or_form="Форма заявки на itr2050.ru",
        official_page="itr2050.ru/komanda.html ; itr2050.ru",
        confidence="высокая",
        note="В базе уже есть сайт; дополнительно отмечен рабочий вход через форму заявки.",
        sources="https://itr2050.ru/komanda.html ; https://itr2050.ru/",
    ),
    ("Краснодарский край", 28): AddedInfo(
        email_or_form="mischenkoi@develug.ru",
        phone="+7 861 279-45-67; +7 861 279-46-32",
        official_page="Карточки ООО СК «Екатеринодар-Сити»",
        confidence="средняя",
        note="Контакты взяты из открытых досье/агрегаторов; лучше перепроверить на официальных документах перед email-рассылкой.",
        sources="https://zachestnyibiznes.ru/company/ul/1092310003389_2310140354_OOO-SK-EKATERINODAR-SITI ; https://companies.rbc.ru/id/1092310003389-ooo-stroitelnaya-kompaniya-ekaterinodar-siti/",
    ),
}


STANDARD_COLUMNS = [
    "Источник",
    "ПРАКТИКУМ",
    "ГОРОД",
    "ФИО",
    "Номер телефона ",
    "Телеграм ",
    "Ответственный а контакт/ кто связывется ",
    "ИНДУСТРИЯ",
    "КОМПАНИЯ",
    "САЙТ",
    "Добавлено: email / форма",
    "Добавлено: телефон",
    "Добавлено: соцсеть / канал",
    "Добавлено: оф. контакты / страница",
    "Добавлено: уверенность",
    "Добавлено: комментарий",
    "Добавлено: источники",
]


def cell_value(ws, row: int, header_map: dict[str, int], token: str) -> str | None:
    for header, column in header_map.items():
        if token in header:
            return ws.cell(row, column).value
    return None


def normalize_header_map(ws) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        value = ws.cell(1, col).value
        if value is not None:
            mapping[str(value).strip()] = col
    return mapping


def should_skip_responsible(value: object) -> bool:
    text = str(value or "").lower()
    return any(name in text for name in EXCLUDED_RESPONSIBLES)


def default_added_info(phone: object, tg: object, site: object) -> AddedInfo:
    has_direct = any(
        str(value).strip() and str(value).strip().lower() != "нет"
        for value in (phone, tg, site)
    )
    if has_direct:
        return AddedInfo(
            confidence="низкая",
            note="В исходной строке уже есть прямой канал; надежного дополнительного публичного контакта без риска ошибки не найдено.",
        )
    return AddedInfo(
        confidence="низкая",
        note="Надежный публичный контакт не найден; нужна ручная проверка ИНН/юрлица или ассистента компании.",
    )


def add_row(ws_out, values: list[object], added_info: AddedInfo) -> None:
    row_index = ws_out.max_row + 1
    row_values = values + [
        added_info.email_or_form,
        added_info.phone,
        added_info.social,
        added_info.official_page,
        added_info.confidence,
        added_info.note,
        added_info.sources,
    ]
    for col_idx, value in enumerate(row_values, start=1):
        cell = ws_out.cell(row=row_index, column=col_idx, value=value)
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        if col_idx > 10 and value:
            cell.font = ADDED_FONT


def autosize(ws) -> None:
    widths = {
        "A": 18,
        "B": 12,
        "C": 18,
        "D": 28,
        "E": 20,
        "F": 22,
        "G": 18,
        "H": 18,
        "I": 52,
        "J": 30,
        "K": 30,
        "L": 22,
        "M": 24,
        "N": 34,
        "O": 14,
        "P": 46,
        "Q": 52,
    }
    for column, width in widths.items():
        ws.column_dimensions[column].width = width


def build_sheet() -> None:
    wb = load_workbook(WORKBOOK_PATH)
    if OUTPUT_SHEET in wb.sheetnames:
        del wb[OUTPUT_SHEET]

    ws_out = wb.create_sheet(title=OUTPUT_SHEET)
    for col_idx, header in enumerate(STANDARD_COLUMNS, start=1):
        cell = ws_out.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    kept_rows = 0
    for sheet_name in TARGET_SHEETS:
        ws = wb[sheet_name]
        header_map = normalize_header_map(ws)
        for row in range(2, ws.max_row + 1):
            practice = cell_value(ws, row, header_map, "ПРАКТИКУМ")
            city = cell_value(ws, row, header_map, "ГОРОД")
            fio = cell_value(ws, row, header_map, "ФИО")
            phone = cell_value(ws, row, header_map, "Номер телефона")
            tg = cell_value(ws, row, header_map, "Телеграм")
            responsible = cell_value(ws, row, header_map, "Ответственный")
            industry = cell_value(ws, row, header_map, "ИНДУСТРИЯ")
            company = cell_value(ws, row, header_map, "КОМПАНИЯ")
            site = cell_value(ws, row, header_map, "САЙТ")

            if not any([practice, city, fio, phone, tg, responsible, industry, company, site]):
                continue
            if should_skip_responsible(responsible):
                continue

            responsible_column = next(
                column for header, column in header_map.items() if "Ответственный" in header
            )
            ws.cell(row, responsible_column, value="Влад")

            added_info = ADDED_CONTACTS.get((sheet_name, row), default_added_info(phone, tg, site))
            add_row(
                ws_out,
                [
                    f"{sheet_name}!{row}",
                    practice,
                    city,
                    fio,
                    phone,
                    tg,
                    "Влад",
                    industry,
                    company,
                    site,
                ],
                added_info,
            )
            kept_rows += 1

    ws_out.freeze_panes = "A2"
    ws_out.auto_filter.ref = ws_out.dimensions
    autosize(ws_out)
    wb.save(WORKBOOK_PATH)
    print(f"Создан лист '{OUTPUT_SHEET}' с {kept_rows} строками.")


if __name__ == "__main__":
    build_sheet()
