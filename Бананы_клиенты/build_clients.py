from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT = "/home/user/poiskklientov/Бананы_клиенты/Клиенты_подписка_на_цены_бананов.xlsx"

DOC = "по документу (раздел 4), в этой сессии не перепроверялось"
CHK2 = "подтверждено двумя независимыми поисками, 25.10.2026"
CHK1 = "найдено одним поиском, 25.10.2026"
CHK_NONE = "не перепроверялось"
DE_NOTE = "Германия: холодный рекламный email запрещён (UWG §7), первое касание через LinkedIn или звонок."
DOLE_SRC = "https://www.doleplc.com/contact"
DOLE_CHK = "по документу: источник - страница контактов Dole plc, страница самой компании не зафиксирована; в этой сессии не перепроверялось"


def r(doc, company, place, role, prio, email, phone, site, source, check, note="", group=""):
    return dict(doc=doc, company=company, place=place, role=role, prio=prio, email=email,
                phone=phone, site=site, source=source, check=check, note=note, group=group)


# ---------- Независимые (лист 1) ----------
indep = [
    # A
    r(1, "Citronex", "Польша, Згожелец", "импортёр, 6 дозревательных центров (PL, RO, HU, BE), ~520 тыс. т бананов в год", "A",
      "info@citronex.pl", "+48 75 772 19 40", "https://citronex.pl", "https://citronex.pl/pl/", DOC),
    r(2, "Targban (бренд Catalina)", "Польша, Плохочин", "импортёр-дозреватель, семейная компания, расширяет дозревалку в Плохочине", "A",
      "hello@targban.com", "+48 22 48 14 800", "https://targban.com", "https://targban.com/en/contact/",
      "по документу; email и телефон подтверждены поиском по targban.com 25.10.2026",
      "Ранее на сайте был старый адрес Poznańska 98 (Ożarów Maz.), сейчас Poznańska 630, Płochocin."),
    r(6, "Spreafico Francesco & F.lli", "Италия, Дольцаго", "импортёр, 4 центра дозревания, бананы >20% выручки", "A",
      "info@spreafico.net", "+39 0341 453650", "https://www.spreafico.net", "https://www.spreafico.net/contatti", DOC),
    r(7, "Orsero (Fruttital)", "Италия", "импорт бананов на своих судах, бренд F.lli Orsero; Simba S.p.A., вероятно, та же группа (дубль не подтверждён)", "A",
      "нет (по документу: через LinkedIn)", "-", "", "", "по документу; в этой сессии не перепроверялось",
      "Своя логистика и аналитика, возможна вторая волна продаж."),
    r(18, "AgroFair Benelux", "Нидерланды, Барендрехт", "импортёр Fairtrade и органических бананов", "A",
      "sales@agrofair.nl, info@agrofair.nl", "+31 180 643 900", "https://www.agrofair.nl", "https://www.agrofair.nl", DOC),
    r(22, "Port International", "Германия, Гамбург", "семейный прямой импортёр, бананы обычные, био, Fairtrade", "A",
      "info@port-international.com", "+49 40 30 10 00 0", "https://port-international.com", "https://port-international.com/impressum/", DOC, DE_NOTE),
    r(23, "Cobana", "Германия, Гамбург", "закупочный союз независимых оптовиков, создан под закупку бананов", "A",
      "нет (сайт блокирует автодоступ)", "+49 40 30305-0", "", "https://www.cio.de/article/3701022/cobana-gmbh-und-co-kg-it-fakten-und-it-strategie.html", DOC,
      "Самая ценная точка в Германии: одна сделка с союзом даёт выход на сеть его оптовиков. Членство отдельных компаний проверять отдельно. " + DE_NOTE),
    r(30, "SHP Group (S.H. Pratt)", "Великобритания, Лутон", "импорт, дозревание, упаковка, дистрибуция", "A",
      "info@shp-group.co.uk", "+44 1582 436500", "https://www.shp-group.co.uk", "https://www.shp-group.co.uk", DOC),
    r(37, "Quiza S.A.", "Польша, Гдыня", "импортёр и дозреватель (контролируемая атмосфера), ~60 тыс. т в год (оценка компании по money.pl, около 2018 г.)", "A",
      "handel@quiza.pl", "+48 58 629 66 00", "https://quiza.pl", "https://quiza.pl/en/kontakt/", CHK2,
      "Юрлицо в источниках то S.A., то Sp. z o.o. - уточнить при оформлении договора."),
    r(38, "Augma UAB", "Литва, Видмантай (Кретингский р-н)", "дозреватель и оптовик, 36 камер дозревания (сайт), называет себя крупнейшим дозревателем бананов в Балтии, филиал в Латвии с 2009 г.", "A",
      "info@augma.lt", "+370 445 75000", "https://augma.lt", "https://augma.lt/en/contacts/", CHK2,
      "Страница контактов примерно двухлетней давности, сверить. Есть отдельное юрлицо Vilniaus Augma (та же группа)."),
    r(39, "Banalat SIA", "Латвия, Рига", "импортёр, дозреватель и оптовик с 1998 г., мощность дозревания от 50 тыс. коробок в неделю (сайт)", "A",
      "info@banalat.eu", "+371 67 367 889", "https://banalat.com", "https://banalat.com/", CHK2,
      "Домен email .eu, сайт на .com. Склад: noliktava@banalat.eu. Телефоны на английской и латышской страницах различаются."),
    r(40, "Consorfrut Polska", "Польша, Краков (офис), Плохочин, Бронише", "оптовик плодоовощей, по сайту с камерами дозревания, склады в Кракове, Плохочине и на Бронише", "A",
      "biuro@consorfrut.pl", "+48 12 342 01 00", "https://www.consorfrut.pl", "https://www.consorfrut.pl/company", CHK2,
      "Не дубль Targban: другой адрес (ul. Wolska 1) и домен. Дозревание заявлено на сайте, моя проверка подтвердила только контакты и адреса."),
    r(41, "Bury Sp. z o.o.", "Польша, Люблин", "оптовик плодоовощей, собственные дозревательные камеры, склады 12 тыс. м² (сайт)", "A",
      "info@bury.pl", "+48 81 469 77 55", "https://bury.pl", "https://bury.pl/kontakt/", CHK2,
      "Номера на старых страницах расходятся, звонить на секретариат и уточнять. Склад бананов: +48 661 611 769 (страница .Trade)."),
    # B
    r(24, "Fruchtimport vanWylick", "Германия, Дюссельдорф", "бананы ~30% выручки, сеть дозревания (74% у Fyffes)", "B",
      "info@vanwylick.de", "+49 221 65032-570", "https://www.vanwylick.de", "https://www.vanwylick.de", DOC, DE_NOTE),
    r(25, "Werder Frucht", "Германия, Гросс-Кройц", "оптовик с дозревательной, член группы Cobana", "B",
      "s.schornberg@vsholding.de", "-", "", "https://www.uni-potsdam.de/de/wirtschaft-transfer-gesellschaft/piw/partner/werder-frucht-gmbh", DOC,
      "Личный адрес сотрудника из публичного каталога партнёров Потсдамского университета: писать адресно и аккуратно. " + DE_NOTE),
    r(42, "Czarpol Sp. z o.o.", "Польша, Гродзиск-Мазовецкий", "дозреватель и оптовик с 1990 г., 30 тыс. коробок в неделю (сайт), точка на рынке Бронише, прямой импорт не подтверждён", "B",
      "czarpol@czarpol.pl", "+48 22 755 68 45", "http://www.czarpol.pl", "http://www.czarpol.pl/", CHK2,
      "Второй телефон: +48 22 73 45 900. Сайт на старом движке."),
    r(43, "Artfoods Sp. z o.o.", "Польша, Краков", "оптовик плодоовощей с дозреванием, ~20 тыс. коробок зелёных бананов в неделю, дозревание 11 тыс. коробок в неделю (сайт, страница старая), прямой импорт не подтверждён", "B",
      "biuro@artfoods.com.pl", "+48 12 341 68 91", "https://artfoods.com.pl", "https://artfoods.com.pl/kontakt/", CHK2,
      "На странице контактов есть личные адреса менеджеров по бананам, в файл не включены."),
    r(44, "NOVEX-BAN (группа NVX)", "Польша, Доманевек-Первы (у Варшавы)", "импортёр (по сайту, прямые поставки из Южной Америки), дозреватель и оптовик, более 500 контейнеров в год (сайт)", "B",
      "biuro@novexban.eu", "-", "https://grupanvx.eu", "https://grupanvx.eu/dojrzewalnia-bananow/", "подтверждено двумя поисками, но оба раз только на сайте группы grupanvx.eu, 25.10.2026",
      "Телефоны на странице личные (владельцев), в файл не включены."),
    r(45, "Fruit Era UAB", "Литва, Вильнюс", "импортёр и дозреватель, 6 камер, 10 тыс. коробок в неделю, до 20 тыс. с партнёрами, поставки в Литву, Латвию и Эстонию (сайт)", "B",
      "info@fruitera.lt", "-", "https://fruitera.lt", "https://fruitera.lt/contacts/", CHK2,
      "Телефоны на странице личные, в файл не включены."),
    r(46, "Karlskroona OÜ", "Эстония, Йыэлехтме (у Таллина)", "оптовик плодоовощей (группа Veikko Laine Oy, Финляндия), своя дозревалка бананов, по сайту единственная специализированная в Эстонии, объём бананов не указан", "B",
      "info@karlskroona.ee", "+372 606 4626", "https://karlskroona.ee", "https://karlskroona.ee/en/contact", CHK1,
      "Заказы: tellimus@karlskroona.ee. Телефон - линия продаж по данным агента."),
    r(47, "Citma UAB", "Литва, Шяуляй", "импортёр, оптовик и дозреватель, 14 камер, 10 тыс. коробок одновременно (сайт)", "B",
      "нет (на сайте контакты только по телефонам)", "+370 41 522251", "https://www.citma.lt", "https://www.citma.lt/en/requisites/", CHK2,
      "Контактные лица с мобильными номерами есть на странице citma.lt/en/contact-persons/, в файл не включены."),
    r(48, "UAB Litbana", "Литва, Вильнюс", "оптовик и импортёр плодоовощей с 1992 г.; по торговым записям Panjiva (данные из выдачи поиска) бананы - топ-товар, поставки из Южной Америки; дозревание не подтверждено", "B",
      "info@litbana.lt", "+370 5 260 19 37", "https://litbana.lt", "https://www.litbana.lt/en/contacts/", CHK2,
      "Данные Panjiva из сниппетов, страницы не открывались."),
    r(49, "SIA Latbana Group", "Латвия, Рига", "по Baltic Times связанная компания Latbana Baltic ввозила эквадорские бананы для сети Maxima Latvija (что это одно юрлицо, не подтверждено), выручка 2024 около 10,7 млн EUR (агрегатор), дозревание и объём бананов не подтверждены", "B (предв.)",
      "нет (сайт не найден)", "+371 67686070", "", "https://www.zl.lv/en/company/latbana-group-sia/", CHK1,
      "В каталоге Lursoft указан другой телефон: +371 67844300."),
    # C
    r(35, "BAMA Gruppen", "Норвегия, Осло", "крупнейший частный дистрибьютор плодоовощей, свои дозревательные", "C",
      "нет (форма на bama.no)", "+47 22 88 05 00", "https://www.bama.no", "https://www.bama.no", DOC),
    r(50, "Baltic Fresh Fruit UAB", "Литва, Laistai (р-н Клайпеды), склад в Вильнюсе", "импортёр плодоовощей широкого ассортимента, бананы лишь часть ассортимента, в вильнюсском складе (3,7 тыс. м²) есть система дозревания бананов (сайт), объём бананов не указан", "C",
      "info@bff.lt", "+370 46 215572", "https://bff.lt", "https://bff.lt/en/contact-us/", CHK2,
      "Закупки цитрусовых и бананов ведёт отдельный сотрудник, его личные контакты в файл не включены."),
    r(51, "A Fruits SIA", "Латвия, Рига", "оптовик фруктов с 2013 г., специализированные системы дозревания бананов (сайт), поставки сетям Латвии, Литвы и Эстонии, объём и число камер не указаны", "C",
      "info@afruits.lv", "+371 675 40 098", "https://afruits.lv", "https://afruits.lv/en/about-us/", CHK2),
    r(52, "Fruit-Expert Poland Sp. z o.o.", "Польша, Забже", "оптовик бананов и дозреватель с 1991 г., 10 камер, склад на 30 тыс. коробок (данные из профиля на job-портале); вероятно, преемник Banalux (тот же адрес в Забже, banalux.pl, biuro@banalux.pl найден одним поиском), по косвенным признакам, не подтверждено", "C",
      "нет (сайт fruit-expert.pl не открылся)", "нет", "", "https://www.aplikuj.pl/pracodawca/347805/fruit-expert-poland-spolka-z-ograniczona-odpowiedzialnoscia", CHK_NONE,
      "Контактов нет, источник подтверждает только существование компании."),
    r(53, "Elbana", "Польша, Краков", "дозреватель, 7 камер давления, 8,4 тыс. коробок, более 15 лет (сайт), продаёт бананы Dole, Chiquita, Del Monte, UTC, собственный импорт не подтверждён", "C",
      "нет (на найденных страницах не показан)", "+48 12 659 08 69", "https://www.elbana.pl", "https://www.elbana.pl/index.php?action=kontakt", CHK1,
      "Второй телефон: +48 12 650 54 50."),
    r(54, "AS Bambona", "Эстония, Таллин", "оптовик-импортёр плодоовощей с 1994 г., продаёт бананы коробками по 18 кг, прямой импорт и дозревание не подтверждены, бананы лишь часть ассортимента", "C",
      "tallinn@bambona.com (ящик склада)", "+372 605 5994", "https://bambona.ee", "https://bambona.ee/kontakt/", CHK2,
      "Есть также tartu@bambona.com. Основной телефон по реестру: +372 605 5666."),
    r(55, "UAB Kauno vaisių ir daržovių prekyba", "Литва, Каунас", "оптовик плодоовощей для сетей Литвы и ЕС (30 тыс. т продукции в год по всем товарам, самозаявление), бананы только в каталоге товаров, импорт и дозревание не подтверждены", "C",
      "info@kvdp.lt", "+370 37 40 00 10", "https://www.kaunovaisiai.lt", "https://www.kaunovaisiai.lt/kontaktai.html", CHK1),
    r(56, "PPUH Polmex Sp. z o.o.", "Польша, Новы-Сонч (точка на Бронише)", "импортёр плодоовощей с 1990 г., бананы в ассортименте на сайте компании, доля бананов, объём и дозревание не найдены", "C",
      "office@polmex.pl", "+48 18 415 67 676", "https://www.polmex.pl", "https://www.polmex.pl/", CHK1,
      "В справочнике Бронише указан ещё bronisze@polmex.pl. На английской странице другой телефон: (18) 442-88-26."),
    r(57, "Fructipol Sp. z o.o.", "Польша, Варшава (Бронише)", "по каталогу bestfoodimporters семейный импортёр и экспортёр плодоовощей, бананы в списке товаров на странице Бронише, импорт бананов и дозревание не подтверждены", "C",
      "нет (сайт fructi.pl по каталогу, не проверен)", "нет", "", "https://www.bronisze.com.pl/sprzedawca/fructipol-sp-z-oo", CHK_NONE,
      "Контактов нет, источник подтверждает только существование компании."),
    r(58, "Fruco Sp. z o.o.", "Польша, Варшава (Бронише)", "по справочнику Бронише многолетний импортёр цитрусовых, бананы в списке товаров, импорт бананов и дозревание не подтверждены", "C",
      "нет (в справочнике только бесплатная почта)", "+48 22 721 60 20", "", "https://www.bronisze.com.pl/sprzedawca/fruco-sp-z-oo", CHK1,
      "Возможна связь с Tropical Trade (одинаковый NIP в сниппетах), не проверено."),
]

# ---------- Группы (лист 2) ----------
groups = [
    r(3, "Hortim International", "Чехия, Брно", "дистрибьютор фруктов", "A", "brno@hortim.cz", "+420 5 4352 2222", "", DOLE_SRC, DOLE_CHK, group="Dole"),
    r(4, "Hortim Prague", "Чехия, Прага", "оптовый рынок Lipence", "A", "praha@hortim.cz", "+420 2 5792 2595", "", DOLE_SRC, DOLE_CHK, group="Dole"),
    r(5, "Hortim-SK", "Словакия, Сенец", "дистрибьютор", "A", "hortim@hortim.sk", "+421 244 872 872", "", DOLE_SRC, DOLE_CHK, group="Dole"),
    r(8, "Dole Italy", "Италия, Кальчо (BG)", "коммерческий офис", "B", "contatto@dole.com", "+39 0363 183541", "", DOLE_SRC, DOLE_CHK, group="Dole"),
    r(9, "Eurobanan Madrid", "Испания, Mercamadrid", "дистрибуция бананов", "A", "madrid@eurobanan.com", "+34 91 779 66 00", "", DOLE_SRC, DOLE_CHK, group="Dole"),
    r(10, "Eurobanan Valencia", "Испания, Mercavalencia", "дистрибуция", "A", "valencia@eurobanan.com", "+34 96 36 73 269", "", DOLE_SRC, DOLE_CHK, group="Dole"),
    r(11, "Eurobanan Alicante", "Испания, Mercalicante", "дистрибуция", "A", "alicante@eurobanan.com", "+34 96 510 77 18", "", DOLE_SRC, DOLE_CHK, group="Dole"),
    r(12, "Eurobanan Murcia", "Испания, Mercamurcia", "дистрибуция", "A", "murcia@eurobanan.com", "+34 96 89 69 226", "", DOLE_SRC, DOLE_CHK, group="Dole"),
    r(13, "Eurobanan Galicia", "Испания, Порриньо", "дистрибуция", "A", "vigo@eurobanan.com", "+34 98 628 80 34", "", DOLE_SRC, DOLE_CHK, group="Dole"),
    r(14, "Eurobanan Lisboa", "Португалия, MARL", "дистрибуция", "A", "escritoriolisboa@eurobanan.com", "+351 219 844 369", "", DOLE_SRC, DOLE_CHK, group="Dole"),
    r(15, "Eurobanan Porto", "Португалия, MAP", "дистрибуция", "A", "escritorioporto@eurobanan.com", "+351 225 370 283", "", DOLE_SRC, DOLE_CHK, group="Dole"),
    r(17, "Anaco & Greeve International", "Нидерланды, Пулдейк", "импорт и дозревание, 40 тыс.+ коробок в неделю, продажи из Антверпена, Флиссингена, Роттердама", "A",
      "info@anacogreeve.nl", "+31 174 638666", "https://www.anacogreeve.nl", DOLE_SRC,
      "по документу; присутствие в списке контактов Dole plc подтверждено поиском 25.10.2026, прямого подтверждения владения Dole не найдено",
      "Принадлежность к группе Dole выведена из страницы контактов Dole plc, прямое подтверждение владения не найдено.", group="Dole (по списку Dole plc)"),
    r(19, "Dole Europe B.V.", "Нидерланды, Блейсвейк", "европейский офис Dole", "B", "нет", "+31 10 244 84 00", "", DOLE_SRC, DOLE_CHK, group="Dole"),
    r(26, "Dole Europe GmbH", "Германия, Гамбург", "штаб Dole в Германии", "B", "kontakt@dole.com", "+49 40 3290 60", "", DOLE_SRC, DOLE_CHK, DE_NOTE, group="Dole"),
    r(28, "Dole France", "Франция, Кавайон", "импорт и дистрибуция", "B", "mail@totalproduce.fr", "+33 432 500 000", "", DOLE_SRC, DOLE_CHK, group="Dole"),
    r(29, "Dole France (Indigo Fruit)", "Франция, Витроль", "импорт и дистрибуция", "B", "sales@indigo-fruit.com", "+33 442 416 680", "", DOLE_SRC, DOLE_CHK, group="Dole"),
    r(32, "AB Banan-Kompaniet", "Швеция, Хельсингборг", "банановая компания в группе Dole", "B", "bk@saba.se", "+46 42 249 600", "", DOLE_SRC, DOLE_CHK, group="Dole"),
    r(33, "Dole Nordic AB", "Швеция, Хельсингборг", "импорт и дистрибуция", "C", "info@dolenordic.com", "+46 42 450 23 00", "", DOLE_SRC, DOLE_CHK, group="Dole"),
    r(34, "Everfresh AB", "Швеция, Хельсингборг", "фрукты и овощи (группа Dole)", "C", "info@everfresh.se", "+46 42 450 23 00", "", DOLE_SRC, DOLE_CHK, group="Dole"),
    r(36, "Dole Nordic A/S", "Дания, Кёге", "импорт и дистрибуция", "C", "нет", "+45 36 15 63 22", "", DOLE_SRC, DOLE_CHK, group="Dole"),
    r(20, "Chiquita Europe B.V.", "Нидерланды, Горинхем", "европейский импорт Chiquita", "B", "нет (по документу: через LinkedIn)", "-", "", "https://panjiva.com/Chiquita-Europe-BV/39680906", "по документу; в этой сессии не перепроверялось", group="Chiquita"),
    r(21, "Greenyard Fresh", "Бельгия, Синт-Катлейне-Ваверс", "называет себя крупнейшим дозревателем бананов в Европе", "B", "нет (скрыт на сайте, форма greenyard.group)", "-", "https://www.greenyard.group", "https://www.greenyard.group/divisions/fresh/local-pages/fresh-hq", DOC, group="Greenyard"),
    r(27, "Compagnie Fruitière", "Франция, Марсель", "~775 тыс. т бананов в год, 22 дозревательных центра в 8 странах", "B", "нет (не подтверждён)", "+33 4 91 10 17 10", "https://compagniefruitiere.fr", "https://compagniefruitiere.fr/wp-content/uploads/2024/12/CIE-FRUITIERE-2024_DEPLIANT_WEB_UK.pdf", DOC, group="Compagnie Fruitière"),
    r(16, "Compagnie Fruitière España", "Испания, Mercamadrid", "дистрибуция бананов и тропиков", "B", "нет (не подтверждён)", "+34 915 076 200", "", "https://proveedores.com/proveedores/compagnie-fruitiere/", DOC, group="Compagnie Fruitière"),
    r(31, "Fyffes", "Великобритания, Ирландия", "крупный импортёр бананов в Европе; во Франции с 01.01.2027 единственным дистрибьютором станет Miles Import (анонс 14.10.2026, три года)", "B",
      "нет (форма на fyffes.com)", "-", "https://www.fyffes.com", "https://www.fyffes.com", DOC,
      "Miles Import (Франция) - отдельный клиент для будущей проверки, в базе его пока нет. Про Miles Import сообщали FreshPlaza, Fruitnet, Retail Times (в выдаче поиска, страницы не открывались).", group="Fyffes"),
]

# ---------- Miles Import (новый клиент, в документе его нет) ----------
miles = r(59, "Miles Import SA", "Франция, Иль-де-Франс (адрес из таможенной записи 2023: Chevilly-Larue, не подтверждён)",
          "французский импортёр и дистрибьютор, основан в 1994 г. (цитрусовые, виноград, манго авиадоставкой), с 2020 г. развивает бренд Fyffes во Франции; с 01.01.2027 станет единственным дистрибьютором Fyffes во Франции по трёхлетнему соглашению (анонс 14.10.2026), бананы и ананасы; наличие собственной дозревалки не подтверждено", "B (предв.)",
          "нет (не найден)", "нет", "", "https://retailtimes.co.uk/fyffes-and-miles-import-announce-strategic-partnership-in-france/",
          "контактов не найдено при поиске 25.10.2026; существование и роль подтверждены сообщениями прессы (в выдаче поиска, страницы не открывались)",
          "Председатель и CEO по заявлению Miles Import в пресс-релизе: Guy Lesvenan. Контакты искать через реестр Infogreffe или SIRENE и LinkedIn. Сообщали также FreshPlaza, Fruitnet, Lebensmittelpraxis.")
_pos = next(i for i, x in enumerate(indep) if x["prio"] == "C")
indep.insert(_pos, miles)

# ---------- Результаты перепроверки 25.10.2026 ----------
DOLE_REV = "перепроверено 25.10.2026: совпадает со страницей контактов Dole plc (по выдаче поиска)"
CHK_REV = "перепроверено 25.10.2026: совпадает с официальной страницей (по выдаче поиска)"
CHK_PART = "перепроверено 25.10.2026: часть данных совпала, см. примечание"
CHK_FAIL = "не подтвердилось при перепроверке 25.10.2026, см. примечание"
CHK_NEW = "найдено при перепроверке 25.10.2026 (один источник)"


def upd(rows, doc, **kw):
    hit = [x for x in rows if x["doc"] == doc]
    assert len(hit) == 1, doc
    hit[0].update(kw)


upd(indep, 1, source="https://citronex.pl/pl/kontakt", check=CHK_REV,
    note="На странице контактов телефон записан как +48 75 77 21 940 (тот же номер), второй номер +48 75 77 21 946. У дочерних структур группы свои телефоны, это головной офис.")
upd(indep, 6, source="https://www.spreafico.net/contatti", check=CHK_REV,
    note="Офис Dolzago. Телефоны рыночных точек: Милан +39 02 59900797, Верона +39 045 8621444.")
upd(indep, 7, email="investor.relations@orserogroup.it (только для инвесторов, коммерческого адреса нет)", phone="+39 0182 560400",
    site="https://www.orserogroup.it", source="https://emarketstorage.it/sites/default/files/comunicati/2022-03/20220318_109835.pdf", check=CHK_NEW,
    note="Найдено при перепроверке в пресс-релизе Orsero за 2022 г. (страница не открывалась): телефон и адрес для связи с инвесторами, актуальность неизвестна. Для продаж нужен личный выход через LinkedIn. Своя логистика и аналитика, возможна вторая волна продаж.")
upd(indep, 18, source="https://www.agrofair.nl/contact-us/", check=CHK_REV,
    note="Адрес: Koopliedenweg 10, 2991 LN Barendrecht.")
upd(indep, 22, source="https://port-international.com/en/legal-notice/", check=CHK_REV,
    note="Страница в поисковом индексе примерно двухлетней давности. Адрес: Lippeltstr. 1, 20097 Hamburg. " + DE_NOTE)
upd(indep, 23, email="info@cobana.com", site="https://www.cobana.com", source="https://www.cobana.com/en/contact/",
    check="найдено при перепроверке 25.10.2026: email и телефон совпали на нескольких страницах cobana.com",
    note="В документе email не был найден (сайт блокировал автодоступ), при перепроверке адрес найден. Адрес: Neue Burg 2, 20457 Hamburg. Самая ценная точка в Германии: одна сделка с союзом даёт выход на сеть его оптовиков. " + DE_NOTE)
upd(indep, 30, email="info@shp-group.co.uk (не подтверждён: на сайте скрыт)", source="https://www.shp-group.co.uk/contact/", check=CHK_PART,
    note="Телефон подтверждён. В отчёте ESG 2022/23 группы указан другой адрес info@shprattgroup.com, проверить на сайте перед рассылкой.")
upd(indep, 24, source="https://www.vanwylick.de/en/contact/", check=CHK_REV,
    note="Адрес: Rolandstr. 44, 40476 Düsseldorf. На сайте есть отдельные адреса для закупок, прямых продаж и торговли бананами. " + DE_NOTE)
upd(indep, 25, email="info@werder-frucht.de", phone="+49 33207 370-0", site="https://www.werder-frucht.de", source="https://www.werder-frucht.de/kontakt",
    check="найдено при перепроверке 25.10.2026 на официальном сайте (в выдаче поиска только этот источник)",
    note="Личный адрес сотрудника из документа заменён общим адресом компании. Для торговли и общепита: gastro@werder-frucht.de. Адрес: Am Frucht- und Frachthof 6, 14550 Groß Kreutz. Член группы Cobana (по странице Потсдамского университета). " + DE_NOTE)
upd(indep, 35, source="https://www.bama.no/kontakt-oss", check=CHK_REV,
    note="Телефон подтверждён, email на сайте не показан, только формы обращения.")

for d in (3, 4, 5, 8, 19, 26, 33, 34, 36):
    upd(groups, d, check=DOLE_REV)
upd(groups, 19, note="Офис в Блейсвейке, email на странице не указан.")
upd(groups, 36, note="На странице два датских адреса: Hasselager +45 86 28 82 22 и Кёге +45 36 15 63 22.")
for d in (9, 10, 11, 12, 13):
    upd(groups, d, source="https://eurobanan.com/contacto/", check=CHK_REV)
upd(groups, 9, note="У Мадрида несколько номеров: +34 91 779 66 00 (страница контактов), 917 796 600 и +34 91 779 66 72 (правовая страница).")
upd(groups, 13, note="На сайте два адреса в районе Виго (рынок Mercado Vigo и Centro Galicia в Порриньо) с одним email.")
for d in (14, 15):
    upd(groups, d, check="email подтверждён на странице Dole plc 25.10.2026, телефон отдельно не подтверждён")
upd(groups, 17, source="https://www.anacogreeve.nl/contact/", check=CHK_REV,
    note="Принадлежность к группе Dole выведена из страницы контактов Dole plc, прямое подтверждение владения не найдено. Актуальный телефон +31 174 638 666 (на старой странице другие номера). Личные адреса специалистов на сайте есть, в файл не включены.")
upd(groups, 28, check=CHK_FAIL,
    note="Адрес mail@totalproduce.fr и офис в Кавайоне не найдены ни на странице Dole plc, ни в других источниках. Total Produce Indigo стала Dole France с 01.01.2023 (по Réussir), головной офис Dole France: 511 Route de la Seds, 13127 Vitrolles. Возможный дубль строки 29.")
upd(groups, 29, check=CHK_FAIL,
    note="sales@indigo-fruit.com не подтвердился. Indigo Fruit входит в Dole France (с 01.01.2023), головной офис в Витроле. Возможный дубль строки 28.")
upd(groups, 32, check=CHK_FAIL,
    note="AB Banan-Kompaniet не найдена ни на странице Dole plc, ни в поиске. Домен saba.se, по сообщению SFVBJ (из выдачи поиска, страница не открывалась), связан с Saba Fresh Cuts, которую Dole продала группе Bama, email мог устареть. Не использовать без проверки.")
upd(groups, 20, check="контактов нет, перепроверять нечего")
upd(groups, 21, phone="+32 15 32 47 00 (Greenyard Fresh Belgium)", source="https://www.greenyard.group/nl/divisions/fresh/local-pages/fresh-belgium", check=CHK_NEW,
    note="Email на сайте скрыт, есть только форма. Головной офис: Strijbroek 10, Sint-Katelijne-Waver, телефон корпоративного офиса +32 15 32 42 00. Адрес info.direct@greenyardfresh.be относится к другому подразделению (Fresh Direct Belgium).")
upd(groups, 27, email="fruitiere@fruitiere.fr", source="https://www.compagniefruitiere.fr/mentions-legales/",
    check="телефон совпал, email найден при перепроверке 25.10.2026 на странице mentions légales",
    note="Адрес: 33 boulevard Ferdinand de Lesseps, 13014 Marseille. В документе email был не подтверждён.")
upd(groups, 16, check="не перепроверялось (страница proveedores.com в поиске не вернулась)")
upd(groups, 31, phone="+353 1 887 2700 (Дублин), +31 10 244 5300 (Роттердам)", source="https://www.fyffes.com/contact-us/", check=CHK_NEW,
    note="Во Франции с 01.01.2027 единственным дистрибьютором станет Miles Import (анонс 14.10.2026, три года), он добавлен на лист «Независимые». Общего email для клиентов на странице контактов нет, только телефоны офисов, страница по данным поиска обновлялась больше полугода назад.")
upd(groups, 31, source="https://www.fyffes.com/contact-us-3/")

HEAD_IND = ["№", "№ в документе", "Компания", "Страна, город", "Роль и объём", "Приор.", "Email", "Телефон", "Сайт",
            "Откуда контакты (страница)", "Проверка контактов", "Примечание"]
HEAD_GRP = ["№", "№ в документе", "Группа", "Компания", "Страна, город", "Роль и объём", "Приор.", "Email", "Телефон", "Сайт",
            "Откуда контакты (страница)", "Проверка контактов", "Примечание"]

FONT = "Arial"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
PRIO_FILL = {"A": "E2EFDA", "B": "FFF2CC", "C": "EDEDED"}


def build_sheet(ws, rows, grouped):
    head = HEAD_GRP if grouped else HEAD_IND
    ws.append(head)
    for c in range(1, len(head) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(name=FONT, bold=True, color="FFFFFF", size=10)
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = BORDER
    ws.row_dimensions[1].height = 32
    for i, x in enumerate(rows, start=1):
        base = [i, x["doc"]]
        if grouped:
            vals = base + [x["group"], x["company"], x["place"], x["role"], x["prio"], x["email"], x["phone"],
                           x["site"], x["source"], x["check"], x["note"]]
            email_col, site_col, src_col, prio_col = 8, 10, 11, 7
        else:
            vals = base + [x["company"], x["place"], x["role"], x["prio"], x["email"], x["phone"],
                           x["site"], x["source"], x["check"], x["note"]]
            email_col, site_col, src_col, prio_col = 7, 9, 10, 6
        ws.append(vals)
        row = ws.max_row
        for c in range(1, len(vals) + 1):
            cell = ws.cell(row=row, column=c)
            cell.font = Font(name=FONT, size=10)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = BORDER
        for col in (site_col, src_col):
            cell = ws.cell(row=row, column=col)
            url = cell.value
            if isinstance(url, str) and url.startswith("http"):
                cell.hyperlink = url
                cell.font = Font(name=FONT, size=10, color="0563C1", underline="single")
        p = str(x["prio"])[:1]
        pc = ws.cell(row=row, column=prio_col)
        pc.fill = PatternFill("solid", fgColor=PRIO_FILL.get(p, "FFFFFF"))
        pc.alignment = Alignment(horizontal="center", vertical="top")
        for col in (1, 2):
            ws.cell(row=row, column=col).alignment = Alignment(horizontal="center", vertical="top")
        ec = ws.cell(row=row, column=email_col)
        if isinstance(ec.value, str) and ec.value.startswith("нет"):
            ec.font = Font(name=FONT, size=10, color="7F7F7F")
    widths_ind = [5, 9, 28, 24, 55, 8, 30, 20, 24, 50, 32, 50]
    widths_grp = [5, 9, 18, 28, 24, 50, 8, 30, 20, 24, 50, 36, 45]
    for idx, w in enumerate(widths_grp if grouped else widths_ind, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = w
    ws.freeze_panes = "D2" if not grouped else "E2"
    ws.auto_filter.ref = ws.dimensions


wb = Workbook()
ws1 = wb.active
ws1.title = "Независимые"
build_sheet(ws1, indep, False)
ws2 = wb.create_sheet("Группы (второй этап)")
build_sheet(ws2, groups, True)

ws3 = wb.create_sheet("Как читать")
legend = [
    ("Что это", "Список компаний, которым можно продавать подписку на сайты с фактическими ценами покупки и продажи бананов (цена подписки $500-1000 в месяц, зависит от компании, по данным заказчика). Состояние на 25.10.2026."),
    ("Откуда список", "Раздел 4 (36 компаний) и раздел 4а (23 компании: Польша, Прибалтика и Miles Import) документа «Рынок бананов Европы: анализ и база импортёров»: https://claude.ai/artifact/61RLiR5jvtqbyENn8PMmXd. Раздел 4 собран в предыдущей сессии, раздел 4а - 25.10.2026."),
    ("Лист «Независимые»", "34 компании: независимые импортёры, дозреватели и оптовики. Продавать в первую очередь, потому что у мультинационалов своя аналитика. Сортировка: приоритет A, B, C."),
    ("Лист «Группы (второй этап)»", "25 записей: 20 компаний, связанных с группой Dole, и по одной-две компании Chiquita, Greenyard, Compagnie Fruitière, Fyffes. По документу идти в них после независимых и предлагать один корпоративный доступ на все страны, массовая рассылка по дочкам рискованна, жалоба одной заблокирует всю группу."),
    ("Приоритет", "A - крупный независимый или ключевая точка входа, B - средний или нужна проверка, C - слабый кандидат: бананы лишь часть ассортимента, мало данных или не подтверждены собственный импорт и дозревание. Приоритеты 37-58 назначены по отчётам агентов, приоритеты 1-36 взяты из документа."),
    ("Колонка «Откуда контакты»", "Страница, на которой найден email или телефон. Для компаний 1-36 при перепроверке 25.10.2026 указана страница контактов компании или группы, где данные совпали с документом (по выдаче поиска). Для компаний 37-58 указана страница из выдачи поиска. Для Orsero указан пресс-релиз 2022 г., для Miles Import - сообщение прессы о соглашении с Fyffes."),
    ("Колонка «Проверка контактов»", "«Подтверждено двумя независимыми поисками» - адрес и телефон дважды найдены в выдаче поиска по домену самой компании (25.10.2026). «Найдено одним поиском» - второго подтверждения нет. «Перепроверено 25.10.2026: совпадает» - контакт из раздела 4 найден в выдаче поиска на официальной странице и совпал с документом. «Найдено при перепроверке» - в документе контакта не было, при перепроверке найден один источник. «Не подтвердилось» - контакт из документа в поиске не найден, использовать только после ручной проверки. Страницы целиком не открывались, потому что прокси блокирует сайты, выдача поиска пересказывается инструментом, поэтому перед рассылкой нужно один раз открыть страницу контактов в браузере."),
    ("Что не включено", "Личные адреса и мобильные телефоны сотрудников (личный адрес Werder Frucht из документа заменён общим адресом компании). Рынок Бронише (Варшава) - оператор площадки, а не клиент. Не проверенные кандидаты (каталог eGospodarka, часть продавцов Бронише, список оптовиков Риги) в файл не вошли."),
    ("Правила рассылки", "Германия: холодный рекламный email запрещён (UWG §7), первое касание через LinkedIn или звонок. Для остальных стран перед рассылкой проверить местные правила (GDPR и национальные нормы) и использовать только общие адреса компаний."),
    ("Известные неточности", "Итог перепроверки 25.10.2026: из 36 компаний раздела 4 контакты полностью совпали у 22, частично у 3 (SHP, Eurobanan Lisboa и Porto), новые контакты найдены у 6 (Orsero, Cobana, Werder Frucht, Greenyard, Compagnie Fruitière, Fyffes), не подтвердились у 3 (Dole France Кавайон, Dole France Витроль, Banan-Kompaniet), не перепроверялись или контактов нет у 2 (Compagnie Fruitière España, Chiquita Europe). Dole France Кавайон и Dole France (Indigo Fruit, Витроль), вероятно, одна компания: Indigo Fruit входит в Dole France с 01.01.2023. Orsero и Simba S.p.A., вероятно, одна группа (не подтверждено). Fruit-Expert Poland, вероятно, преемник Banalux (не подтверждено). Принадлежность Anaco & Greeve к Dole выведена из списка контактов Dole plc, прямого подтверждения владения нет. Сайты Royal Banane и Di Manno не открывались, в список не включены; Italfrutta Di Manno найдена в старой статье FreshPlaza (2018), актуальность неизвестна."),
]
ws3.append(["Пункт", "Описание"])
for c in (1, 2):
    cell = ws3.cell(row=1, column=c)
    cell.font = Font(name=FONT, bold=True, color="FFFFFF", size=10)
    cell.fill = HEAD_FILL
    cell.alignment = Alignment(horizontal="center", vertical="center")
for k, v in legend:
    ws3.append([k, v])
    row = ws3.max_row
    ws3.cell(row=row, column=1).font = Font(name=FONT, bold=True, size=10)
    ws3.cell(row=row, column=2).font = Font(name=FONT, size=10)
    for c in (1, 2):
        ws3.cell(row=row, column=c).alignment = Alignment(wrap_text=True, vertical="top")
        ws3.cell(row=row, column=c).border = BORDER
ws3.column_dimensions["A"].width = 30
ws3.column_dimensions["B"].width = 120

wb.save(OUT)
print("saved", OUT, len(indep), len(groups))
