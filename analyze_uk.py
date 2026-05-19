#!/usr/bin/env python3
"""
Анализ рентабельности по УК за 2025 год.
Компания: Сектор Безопасности / Безопасный город
"""

import pandas as pd
import re
import warnings
warnings.filterwarnings('ignore')

# ============================================================
# 0. КОНФИГУРАЦИЯ
# ============================================================

FILES = {
    'addr': '1. Адресная программа на 2025.xlsx',
    'income': '2025 - Оплата за Домофон и ВН xlsx..xlsx',
    'internet': '2025 - Оплата интернета по объектам.xlsx',
    'registry': '2025_Реестр_счетов_на_оплату_и_учет_расходов.xlsx',
    'engineers': 'СЗ Сервис партнеры 2025.xlsx',
}

# Категории расходов из реестра, которые ИСКЛЮЧАЕМ
EXCLUDE_CATEGORIES = [
    'спутник долг',
    'займ спутник',
    'займ спутник ???',
    'спутник командир',
    'расх марат',       # отдельный анализ
    'овердрафт',
    'проценты',
    'штрафы',
    '???',
]

# Категории — это ФОТ (обрабатываем отдельно)
FOT_CATEGORIES = ['фот', 'фот первичка', 'фот вторичка', 'зп вторичка', 'фот эртх']

# Нужно уточнить у пользователя
PENDING_CATEGORIES = [
    'спутник транзит',
    'транзит спутник',
    'транзитные спутник',
    'расход спутник',
]


# ============================================================
# 1. НОРМАЛИЗАЦИЯ НАЗВАНИЙ УК
# ============================================================

UK_MAP = {
    # --- Адресная программа ---
    'ООО "УК"ЗАПАД"': 'Запад',
    'ООО "Надежное упрвление"': 'Надежное управление',
    'ООО "УК Аурика"': 'Аурика',
    'ООО "УК Суходолье"': 'Суходолье',
    'ООО "УК Линейная"': 'Линейная',
    'ООО "УК ТЭСК"': 'ТЭСК',
    'ООО "УК на Пражской (5415)"': 'На Пражской',
    'ООО "УК на Ткацком"': 'На Ткацком',
    'ООО "УК"ЖЭУ №9"': 'ЖЭУ 9',
    'ООО ИРИДА': 'Ирида',
    'ООО Монолит': 'Монолит',
    'ООО УК Альтернатива': 'Альтернатива',
    'ООО УК Восток-сити': 'Восток-сити',
    'ООО УК Ладья': 'Ладья',
    'ООО УК Тобол': 'Тобол',
    'ООО УК Усадьба': 'Усадьба',
    'ООО УК на Полевой': 'На Полевой',
    'ООО Юг': 'Юг',
    'ООО "ЖСУ"': 'ЖСУ',
    'УК Алсу': 'Алсу',
    'УК Альметьевск': 'Альметьевск',
    'УК Новация': 'Новация',
    'УК Московского района': 'ЖК Моск.района',
    'УК Радужный': 'Радужный',
    'УК Высота': 'Высота',
    'Жилкомплекс': 'Жилкомплекс',
    'Салават Купере': 'Салават Купере',
    'ТСЖ 6': 'ТСЖ №6',
    'ТСЖ Домуправ': 'ТСЖ Домуправ',
    'ТСЖ Современник': 'Современник',
    'ТСЖ Фукса 12': 'Фукса 12',
    'ТСЖ Электрон': 'Электрон',
    'тсн электрон': 'Электрон',
    'ТСН Вавилон': 'Вавилон',
    'тсн вавилон': 'Вавилон',
    'Тринити': 'Тринити',
    'ТЭСК': 'ТЭСК',
    'ЖКХ Московского рна': 'ЖК Моск.района',
    'ЖКХ Московского рна ': 'ЖК Моск.района',
    'Алсу': 'Алсу',
    'Радужный': 'Радужный',
    'Запад': 'Запад',
    'ЖСУ': 'ЖСУ',
    'жсу': 'ЖСУ',
    'Восток-сити': 'Восток-сити',
    'Юг': 'Юг',
    'Усадьба': 'Усадьба',
    'УК Монолит': 'Монолит',
    'УК На Пражской': 'На Пражской',
    'УК на Полевой': 'На Полевой',
    'УК Альтернатива': 'Альтернатива',
    'Альтернатива': 'Альтернатива',
    'иннофонд': 'Жилищный фонд',
    'тринити': 'Тринити',
    'ЖЭУ Осиново': 'ЖЭУ Осиново',
    'ЭРТХА': 'ЭРТХ',
    'ЭРТХ': 'ЭРТХ',
    'ИК ПГТ Васильево': 'ИК ПГТ Васильево',
    'ТСЖ Волжанка': 'ТСЖ Волжанка',
    'ТСЖ ЖСК №9': 'ТСЖ ЖСК №9',
    'ТСЖ Кедр': 'ТСЖ Кедр',
    'ТСЖ №7': 'ТСЖ №7',
    'Застройщик': 'Застройщик',
    'застройщик': 'Застройщик',
    # --- Доходы (контроль оплаты) ---
    'Жилкомплекс ООО': 'Жилкомплекс',
    'ТСЖ № 6': 'ТСЖ №6',
    'Салават Купере УК ООО': 'Салават Купере',
    'Запад УК ООО': 'Запад',
    'Надежное управление ООО': 'Надежное управление',
    'Линейная УК ООО': 'Линейная',
    'Суходолье УК ООО': 'Суходолье',
    'Аурика (Гольцова 10)': 'Аурика',
    'ЖЭУ № 9 через УК ОО': 'ЖЭУ 9',
    'ЖЭУ № 9 ООО через ТРИЦ': 'ЖЭУ 9',
    'УК Ладья ООО': 'Ладья',
    'УК на Пражской ООО': 'На Пражской',
    'УК Высота': 'Высота',
    'УК на Ткацком': 'На Ткацком',
    'Восток Сити (Широтная 125)': 'Восток-сити',
    'УК ИРИДА': 'Ирида',
    'УК Тобол (Локомотивная 116)': 'Тобол',
    'УК Усадьба (Червишевский тракт 94)': 'Усадьба',
    'УК на Полевой (Черныш 2бк2)': 'На Полевой',
    'УК МОНОЛИТ': 'Монолит',
    'ТЭСК ООО': 'ТЭСК',
    'УК ЮГ (Одесская 44)': 'Юг',
    'Альметьевск УК ООО': 'Альметьевск',
    'Новация УК ООО': 'Новация',
    'Алсу (ЕРЦ общий) УК ООО (остатки от спутника)': 'Алсу',
    'Алсу (ЕРЦ общий) УК ООО': 'Алсу',
    'УК ЖКХ Московского района ООО': 'ЖК Моск.района',
    'УК ЖКХ Московского района (Декабристов 178а)': 'ЖК Моск.района',
    'ЭР-Телеком Холдинг ООО': 'ЭРТХ',
    'ТСЖ Фукса 12': 'Фукса 12',
    'Тринити ООО': 'Тринити',
    'ЗИОН ООО': 'ЗИОН',
    'УК Современник (С. Хакима 60)': 'Современник',
    'Иннокомфорт': 'Иннокомфорт',
    'УК Жилищный фонд ООО': 'Жилищный фонд',
    'Жилсервисуют ООО': 'ЖСУ',
    'Татэнергосбыт': None,
    'поступило от Спутника СБ': None,
    'поступило от Спутника БГ': None,
    'Прочие доходы(по расч. счету)': None,
    'займы от Марата': None,
    # --- Интернет ---
    'Альмет': 'Альметьевск',
    'Альметьевск': 'Альметьевск',
    'запад': 'Запад',
    'ирида': 'Ирида',
    'юг': 'Юг',
    'ГЭ': None,
    'ЖКХ': 'ЖК Моск.района',
    'Жэу 9': 'ЖЭУ 9',
    'ЗИОН': 'ЗИОН',
    'нет статики': None,
}

# Исключаемые из доходов
INCOME_EXCLUDE = [
    'Татэнергосбыт',
    'поступило от Спутника СБ',
    'поступило от Спутника БГ',
    'Прочие доходы(по расч. счету)',
    'займы от Марата',
]


def normalize_uk(name):
    """Нормализация названия УК."""
    if pd.isna(name):
        return None
    name = str(name).strip()
    if name in UK_MAP:
        return UK_MAP[name]
    # Пробуем частичное совпадение
    for k, v in UK_MAP.items():
        if k in name or name in k:
            return v
    return name if name and name != 'nan' else None


# ============================================================
# 2. АДРЕСНАЯ ПРОГРАММА → подъезды, камеры, спутники по УК
# ============================================================

def parse_address_program():
    """Парсит адресную программу, возвращает df: uk, podezdy, cameras, sputnik."""
    f = FILES['addr']
    rows = []

    for sheet in pd.ExcelFile(f).sheet_names:
        df = pd.read_excel(f, sheet_name=sheet, header=None)
        if sheet == 'ВИДЕОКАМЕРЫ':
            continue  # handled separately

        # Находим строку заголовка
        hdr = None
        for i in range(min(5, len(df))):
            vals = [str(df.iloc[i, j]).strip().lower() for j in range(len(df.columns))]
            if any('улица' in v or 'адрес' in v or v == 'ук' for v in vals):
                hdr = i
                break
        if hdr is None:
            continue

        headers = [str(df.iloc[hdr, j]).strip() for j in range(len(df.columns))]

        def find_col(names):
            for j, h in enumerate(headers):
                for n in names:
                    if n.lower() in h.lower():
                        return j
            return None

        street_c = find_col(['улица', 'адрес'])
        house_c = find_col(['дом', '№ дома'])
        pod_c = find_col(['подъезд', 'под-в', 'под.'])
        uk_c = find_col(['ук'])
        sput_c = find_col(['спутник'])
        cam_c = find_col(['камер'])

        for i in range(hdr + 1, len(df)):
            r = df.iloc[i]
            uk = normalize_uk(r[uk_c]) if uk_c is not None else None
            if not uk:
                continue

            pod = 0
            if pod_c is not None and pd.notna(r[pod_c]):
                try:
                    pod = int(float(str(r[pod_c]).split(',')[0].strip()))
                except:
                    pod = 0

            sput = 0
            if sput_c is not None and pd.notna(r[sput_c]):
                sv = str(r[sput_c]).strip()
                if sv and sv not in ['', 'nan', '0']:
                    try:
                        sput = int(float(sv))
                    except:
                        sput = 1

            cam = 0
            if cam_c is not None and pd.notna(r[cam_c]):
                cv = str(r[cam_c]).strip()
                if '+' in cv or 'да' in cv.lower():
                    cam = 1
                elif cv and cv not in ['', 'nan', '0']:
                    m = re.search(r'\d+', cv)
                    cam = int(m.group()) if m else 0

            if pod > 0 or sput > 0 or cam > 0:
                rows.append({'uk': uk, 'podezdy': pod, 'cameras': cam, 'sputnik': sput})

    # ВИДЕОКАМЕРЫ
    df_cam = pd.read_excel(f, sheet_name='ВИДЕОКАМЕРЫ', header=0)
    for i in range(len(df_cam)):
        r = df_cam.iloc[i]
        uk = normalize_uk(r.iloc[4]) if pd.notna(r.iloc[4]) else None
        if not uk:
            continue
        cv = str(r.iloc[3]).strip() if pd.notna(r.iloc[3]) else '0'
        m = re.search(r'\d+', cv)
        cam = int(m.group()) if m else 0
        if cam > 0:
            rows.append({'uk': uk, 'podezdy': 0, 'cameras': cam, 'sputnik': 0})

    df_addr = pd.DataFrame(rows)
    return df_addr.groupby('uk').agg(
        podezdy=('podezdy', 'sum'),
        cameras=('cameras', 'sum'),
        sputnik=('sputnik', 'sum'),
    ).reset_index()


# ============================================================
# 2b. АБОНЕНТЫ ПО ГОРОДАМ
# ============================================================

def parse_subscribers():
    """Считает абонентов по городам: домофон (из адресных листов) и ВН (из ВИДЕОКАМЕРЫ)."""
    f = FILES['addr']
    domofon = {}
    vn = {}

    # --- Домофон: кол-во кв из адресных листов ---
    city_sheets = {
        'Тюмень': 'Тюмень',
        'Альметьевск': ' Альметьевск',
        'Казань': 'Казань',
    }
    for city, sheet in city_sheets.items():
        df = pd.read_excel(f, sheet_name=sheet, header=0)
        kv_col = [c for c in df.columns if 'кол' in str(c).lower() and 'кв' in str(c).lower()]
        if kv_col:
            domofon[city] = int(pd.to_numeric(df[kv_col[0]], errors='coerce').sum())

    # Зеленодольск + Осиново + Васильево + Нижние Вязовые
    zd_total = 0
    for sheet in ['Зеленодольск']:
        df = pd.read_excel(f, sheet_name=sheet, header=0)
        kv_col = [c for c in df.columns if 'кол' in str(c).lower() and 'кв' in str(c).lower()]
        if kv_col:
            zd_total += int(pd.to_numeric(df[kv_col[0]], errors='coerce').sum())

    for sheet in ['Поселок Осиново', 'Поселок Васильево', 'Нижние Вязовые']:
        df = pd.read_excel(f, sheet_name=sheet, header=None)
        for i, row in df.iterrows():
            vals = [str(v).lower() for v in row.values if pd.notna(v)]
            if any('кол' in v and 'кв' in v for v in vals):
                df2 = pd.read_excel(f, sheet_name=sheet, header=i)
                kv_col = [c for c in df2.columns if 'кв' in str(c).lower() and 'номер' not in str(c).lower()]
                if kv_col:
                    zd_total += int(pd.to_numeric(df2[kv_col[0]], errors='coerce').sum())
                break
    domofon['Зеленодольск+О+В+НВ'] = zd_total

    # Иннополис — из строки "Аб-ты" (ИТОГО по УК)
    df = pd.read_excel(f, sheet_name='Иннополис', header=None)
    for i, row in df.iterrows():
        vals = [str(v).lower() for v in row.values if pd.notna(v)]
        if any('название' in v for v in vals):
            df2 = pd.read_excel(f, sheet_name='Иннополис', header=i)
            domofon['Иннополис'] = int(pd.to_numeric(df2['Аб-ты'], errors='coerce').dropna().iloc[:-1].sum())
            break

    # ЗИОН — только домофон (включая "домофон/камеры", исключая "камеры" и "домофон/")
    df = pd.read_excel(f, sheet_name='ЗИОН', header=None)
    for i, row in df.iterrows():
        vals = [str(v).lower() for v in row.values if pd.notna(v)]
        if any('кв' in v for v in vals):
            df2 = pd.read_excel(f, sheet_name='ЗИОН', header=i)
            col6 = df2.iloc[:, 6].astype(str).str.strip()
            mask = col6.str.contains('домофон', na=False) & (col6 != 'домофон/')
            domofon['ЗИОН'] = int(pd.to_numeric(df2.loc[mask, 'кол-во кв'], errors='coerce').sum())
            break

    # Ручные корректировки (данные не в адресной программе)
    domofon['Казань'] = domofon.get('Казань', 0) + 18724     # ЭРТХ
    domofon['Иннополис'] = domofon.get('Иннополис', 0) + 85  # ушли в 2026, но в 2025 считаем
    domofon['ЗИОН'] = 78  # в файле данные изменились, эталон 78

    # --- ВН: из ВИДЕОКАМЕРЫ (колонка Квартиры), дедупликация по адресу ---
    df_cam = pd.read_excel(f, sheet_name='ВИДЕОКАМЕРЫ', header=0)
    df_cam['kv'] = pd.to_numeric(df_cam['Квартиры'], errors='coerce')
    df_cam['addr_key'] = (
        df_cam['город'].astype(str).str.strip().str.lower() + '|' +
        df_cam['улица'].astype(str).str.strip().str.lower() + '|' +
        df_cam['дом'].astype(str).str.strip().str.lower()
    )
    df_cam_dedup = df_cam.drop_duplicates(subset='addr_key', keep='first')
    for i, row in df_cam_dedup.iterrows():
        city = str(row['город']).strip()
        kv = row['kv']
        if pd.notna(kv) and kv > 0:
            if city.lower() == 'зион':
                city = 'ЗИОН'
            vn[city] = vn.get(city, 0) + int(kv)

    # Зеленодольск ВН (ключ в ВИДЕОКАМЕРЫ = "Зеленодольск", а не "Зеленодольск+О+В+НВ")
    vn['Зеленодольск+О+В+НВ'] = vn.get('Зеленодольск', 246)

    # Собираем таблицу
    cities_order = ['Тюмень', 'Альметьевск', 'Зеленодольск+О+В+НВ', 'Казань', 'Иннополис', 'ЗИОН']
    rows = []
    for city in cities_order:
        d = domofon.get(city, 0)
        v = vn.get(city, 0)
        rows.append({'Город': city, 'Домофон': d, 'ВН': v, 'ИТОГО': d + v})

    df_sub = pd.DataFrame(rows)
    total = pd.DataFrame([{
        'Город': 'ИТОГО',
        'Домофон': df_sub['Домофон'].sum(),
        'ВН': df_sub['ВН'].sum(),
        'ИТОГО': df_sub['ИТОГО'].sum(),
    }])
    df_sub = pd.concat([df_sub, total], ignore_index=True)
    return df_sub


def parse_detailed_subscribers():
    """Детальная таблица: город → УК → абоненты, подъезды, инженеры."""
    f = FILES['addr']
    rows = []

    # --- ТЮМЕНЬ ---
    df = pd.read_excel(f, sheet_name='Тюмень', header=0)
    df['под'] = pd.to_numeric(df['Всего подъездов'], errors='coerce').fillna(0)
    df['кв'] = pd.to_numeric(df['кол-во кв '], errors='coerce').fillna(0)
    for uk, grp in df.groupby('УК'):
        uk_norm = normalize_uk(uk)
        кв = int(grp['кв'].sum())
        под = int(grp['под'].sum())
        инженеры = ', '.join(sorted(grp['ФИО'].dropna().unique()))
        rows.append({'Город': 'Тюмень', 'УК': uk_norm, 'Домофон': кв, 'ВН': 0, 'Подъезды': под, 'Инженеры': инженеры})

    # ВН Тюмень из ВИДЕОКАМЕРЫ (суммируем, а не перезаписываем)
    df_cam = pd.read_excel(f, sheet_name='ВИДЕОКАМЕРЫ', header=0)
    df_cam['kv'] = pd.to_numeric(df_cam['Квартиры'], errors='coerce').fillna(0)
    for _, r in df_cam[df_cam['город'].str.strip() == 'Тюмень'].iterrows():
        uk_norm = normalize_uk(r['УК'])
        kv = int(r['kv'])
        if kv > 0:
            found = False
            for row in rows:
                if row['Город'] == 'Тюмень' and row['УК'] == uk_norm:
                    row['ВН'] += kv
                    found = True
                    break
            if not found:
                rows.append({'Город': 'Тюмень', 'УК': uk_norm, 'Домофон': 0, 'ВН': kv, 'Подъезды': 0, 'Инженеры': ''})

    # ТСН Электрон (нет в адресной) — добавляем ВН=282 к существующей записи
    for row in rows:
        if row['Город'] == 'Тюмень' and row['УК'] == 'Электрон':
            row['ВН'] = 282
            break
    else:
        rows.append({'Город': 'Тюмень', 'УК': 'Электрон', 'Домофон': 282, 'ВН': 282, 'Подъезды': 0, 'Инженеры': 'Щекалев Артем'})

    # ТСН Вавилон (ВН из ВИДЕОКАМЕРЫ, домофон = 0)
    found = any(r['Город'] == 'Тюмень' and r['УК'] == 'Вавилон' for r in rows)
    if not found:
        rows.append({'Город': 'Тюмень', 'УК': 'Вавилон', 'Домофон': 0, 'ВН': 54, 'Подъезды': 0, 'Инженеры': 'Балашов Виталий'})

    # --- АЛЬМЕТЬЕВСК ---
    df = pd.read_excel(f, sheet_name=' Альметьевск', header=0)
    df['под'] = pd.to_numeric(df['ослуживаемые подъезды'], errors='coerce').fillna(0)
    df['кв'] = pd.to_numeric(df['кол-во кв'], errors='coerce').fillna(0)
    for uk, grp in df.groupby('УК'):
        uk_norm = normalize_uk(uk)
        кв = int(grp['кв'].sum())
        под = int(grp['под'].sum())
        rows.append({'Город': 'Альметьевск', 'УК': uk_norm, 'Домофон': кв, 'ВН': 0, 'Подъезды': под, 'Инженеры': 'Касимов А.'})

    # ВН Альметьевск (суммируем)
    for _, r in df_cam[df_cam['город'].str.strip() == 'Альметьевск'].iterrows():
        uk_norm = normalize_uk(r['УК'])
        kv = int(r['kv'])
        if kv > 0:
            for row in rows:
                if row['Город'] == 'Альметьевск' and row['УК'] == uk_norm:
                    row['ВН'] += kv
                    break

    # --- ЗЕЛЕНОДОЛЬСК+О+В+НВ ---
    df = pd.read_excel(f, sheet_name='Зеленодольск', header=0)
    df['под'] = pd.to_numeric(df['кол-во под-в'], errors='coerce').fillna(0)
    df['кв'] = pd.to_numeric(df['кол-во кв'], errors='coerce').fillna(0)
    for uk, grp in df.groupby('УК'):
        uk_norm = normalize_uk(uk)
        кв = int(grp['кв'].sum())
        под = int(grp['под'].sum())
        инженеры = ', '.join(sorted(grp['ФИО'].dropna().unique())) if 'ФИО' in grp.columns else 'Калачев Артем'
        rows.append({'Город': 'ЗД+О+В+НВ', 'УК': uk_norm, 'Домофон': кв, 'ВН': 0, 'Подъезды': под, 'Инженеры': инженеры or 'Калачев Артем'})

    # Осиново, Васильево, Нижние Вязовые
    for sheet, city_label in [('Поселок Осиново', 'ЗД+О+В+НВ'), ('Поселок Васильево', 'ЗД+О+В+НВ'), ('Нижние Вязовые', 'ЗД+О+В+НВ')]:
        df = pd.read_excel(f, sheet_name=sheet, header=None)
        for i, row_data in df.iterrows():
            vals = [str(v).lower() for v in row_data.values if pd.notna(v)]
            if any('кол' in v and 'кв' in v for v in vals):
                df2 = pd.read_excel(f, sheet_name=sheet, header=i)
                kv_col = [c for c in df2.columns if 'кв' in str(c).lower() and 'номер' not in str(c).lower()]
                uk_col = [c for c in df2.columns if 'ук' in str(c).lower()]
                pod_col = [c for c in df2.columns if 'под' in str(c).lower()]
                fio_col = [c for c in df2.columns if 'фио' in str(c).lower()]
                if kv_col and uk_col:
                    for _, r in df2.iterrows():
                        uk = normalize_uk(r[uk_col[0]])
                        if not uk:
                            continue
                        kv_val = pd.to_numeric(r[kv_col[0]], errors='coerce')
                        кв = int(kv_val) if pd.notna(kv_val) else 0
                        if pod_col:
                            pod_val = pd.to_numeric(r[pod_col[0]], errors='coerce')
                            под = int(pod_val) if pd.notna(pod_val) else 0
                        else:
                            под = 0
                        инж = str(r[fio_col[0]]) if fio_col and pd.notna(r[fio_col[0]]) else 'Калачев Артем'
                        if инж == 'nan':
                            инж = 'Калачев Артем'
                        # Добавляем к существующей УК или создаём новую
                        found = False
                        for row in rows:
                            if row['Город'] == city_label and row['УК'] == uk:
                                row['Домофон'] += кв
                                row['Подъезды'] += под
                                found = True
                                break
                        if not found:
                            rows.append({'Город': city_label, 'УК': uk, 'Домофон': кв, 'ВН': 0, 'Подъезды': под, 'Инженеры': инж})
                break

    # ВН Зеленодольск (суммируем)
    for _, r in df_cam[df_cam['город'].str.strip() == 'Зеленодольск'].iterrows():
        uk_norm = normalize_uk(r['УК'])
        kv = int(r['kv'])
        if kv > 0:
            for row in rows:
                if row['Город'] == 'ЗД+О+В+НВ' and row['УК'] == uk_norm:
                    row['ВН'] += kv
                    break

    # --- КАЗАНЬ ---
    df = pd.read_excel(f, sheet_name='Казань', header=0)
    df['под'] = pd.to_numeric(df['кол-во под-в'], errors='coerce').fillna(0)
    df['кв'] = pd.to_numeric(df['кол-во кв'], errors='coerce').fillna(0)
    for uk, grp in df.groupby('УК'):
        uk_norm = normalize_uk(uk)
        кв = int(grp['кв'].sum())
        под = int(grp['под'].sum())
        rows.append({'Город': 'Казань', 'УК': uk_norm, 'Домофон': кв, 'ВН': 0, 'Подъезды': под, 'Инженеры': 'Садриев Ильдус'})

    # ЭРТХ Казань
    rows.append({'Город': 'Казань', 'УК': 'ЭРТХ', 'Домофон': 18724, 'ВН': 0, 'Подъезды': 0, 'Инженеры': ''})

    # ВН Казань (из ВИДЕОКАМЕРЫ, дедупликация)
    kazan_cam = df_cam[df_cam['город'].str.strip() == 'Казань'].copy()
    kazan_cam['addr_key'] = (
        kazan_cam['улица'].astype(str).str.strip().str.lower() + '|' +
        kazan_cam['дом'].astype(str).str.strip().str.lower()
    )
    kazan_cam = kazan_cam.drop_duplicates(subset='addr_key', keep='first')
    vn_kazan_total = int(kazan_cam['kv'].sum())
    # ВН Казань привязана к УК Московского района
    for row in rows:
        if row['Город'] == 'Казань' and row['УК'] == 'ЖК Моск.района':
            row['ВН'] = vn_kazan_total
            break

    # --- ИННОПОЛИС ---
    rows.append({'Город': 'Иннополис', 'УК': 'Жилищный фонд', 'Домофон': 1115, 'ВН': 0, 'Подъезды': 42, 'Инженеры': ''})
    rows.append({'Город': 'Иннополис', 'УК': 'Иннокомфорт', 'Домофон': 105, 'ВН': 0, 'Подъезды': 4, 'Инженеры': ''})

    # --- ЗИОН ---
    rows.append({'Город': 'ЗИОН', 'УК': 'Тринити', 'Домофон': 78, 'ВН': 92, 'Подъезды': 15, 'Инженеры': ''})
    rows.append({'Город': 'ЗИОН', 'УК': 'Застройщик', 'Домофон': 0, 'ВН': 440, 'Подъезды': 0, 'Инженеры': ''})

    df_detail = pd.DataFrame(rows)

    # Сортируем: по городу, затем по домофон (убывание)
    city_order = ['Тюмень', 'Альметьевск', 'ЗД+О+В+НВ', 'Казань', 'Иннополис', 'ЗИОН']
    df_detail['city_order'] = df_detail['Город'].map({c: i for i, c in enumerate(city_order)})
    df_detail = df_detail.sort_values(['city_order', 'Домофон'], ascending=[True, False])
    df_detail = df_detail.drop(columns='city_order').reset_index(drop=True)

    return df_detail


def parse_tyumen_engineers():
    """Детальная разбивка по Тюмени: УК → подъезды по инженерам + дома Балашова по ВН."""
    f = FILES['addr']
    df = pd.read_excel(f, sheet_name='Тюмень', header=0)
    df['под'] = pd.to_numeric(df['Всего подъездов'], errors='coerce').fillna(0)
    df['кв'] = pd.to_numeric(df['кол-во кв '], errors='coerce').fillna(0)

    # ВН из ВИДЕОКАМЕРЫ
    df_cam = pd.read_excel(f, sheet_name='ВИДЕОКАМЕРЫ', header=0)
    df_cam['kv'] = pd.to_numeric(df_cam['Квартиры'], errors='coerce').fillna(0)
    vn_by_uk = {}
    for _, r in df_cam[df_cam['город'].str.strip() == 'Тюмень'].iterrows():
        uk_norm = normalize_uk(r['УК'])
        kv = int(r['kv'])
        if kv > 0:
            vn_by_uk[uk_norm] = vn_by_uk.get(uk_norm, 0) + kv

    # Дома Балашова по ВН (от пользователя)
    balashov_houses = {
        'ЖСУ': 2, 'На Пражской': 1, 'ТЭСК': 1, 'Запад': 8,
        'Юг': 2, 'Ирида': 3, 'Восток-сити': 1, 'На Полевой': 1,
        'Монолит': 8, 'Тобол': 1, 'Электрон': 1, 'Вавилон': 1,
    }

    # Собираем данные по УК
    rows = []
    for uk, grp in df.groupby('УК'):
        uk_norm = normalize_uk(uk)
        кв = int(grp['кв'].sum())
        под_total = int(grp['под'].sum())
        vn = vn_by_uk.get(uk_norm, 0)

        # Подъезды по инженерам
        eng_pod = {}
        for eng, eng_grp in grp.groupby('ФИО'):
            if pd.notna(eng):
                eng_pod[str(eng)] = int(eng_grp['под'].sum())

        row = {
            'УК': uk_norm,
            'Аб домофон': кв,
            'Аб ВН': vn,
            'Подъезды': под_total,
            'Щекалев (под)': eng_pod.get('Щекалев Артем', 0),
            'Воронцов (под)': eng_pod.get('Воронцов Денис', 0),
            'Полуянов (под)': eng_pod.get('Полуянов Игорь', 0),
            'Агинов (под)': eng_pod.get('Агинов', 0),
            'Балашов (дома ВН)': balashov_houses.get(uk_norm, 0),
        }
        rows.append(row)

    # Добавляем УК из ВИДЕОКАМЕРЫ, которых нет в адресной программе
    for uk_norm, vn in vn_by_uk.items():
        if not any(r['УК'] == uk_norm for r in rows):
            rows.append({
                'УК': uk_norm,
                'Аб домофон': 0,
                'Аб ВН': vn,
                'Подъезды': 0,
                'Щекалев (под)': 0,
                'Воронцов (под)': 0,
                'Полуянов (под)': 0,
                'Агинов (под)': 0,
                'Балашов (дома ВН)': balashov_houses.get(uk_norm, 0),
            })

    # ТСН Электрон (ВН=282 из ВИДЕОКАМЕРЫ, в адресной = "тсн электрон")
    for row in rows:
        if row['УК'] == 'Электрон':
            row['Аб ВН'] = 282
            row['Балашов (дома ВН)'] = 1
            break

    # Вавилон (нет в адресной, ВН=54 из ВИДЕОКАМЕРЫ)
    found = any(r['УК'] == 'Вавилон' for r in rows)
    if not found:
        rows.append({
            'УК': 'Вавилон',
            'Аб домофон': 0,
            'Аб ВН': 54,
            'Подъезды': 0,
            'Щекалев (под)': 0,
            'Воронцов (под)': 0,
            'Полуянов (под)': 0,
            'Агинов (под)': 0,
            'Балашов (дома ВН)': 1,
        })

    # Ручные корректировки
    for row in rows:
        # Запад: Воронцов=184, Полуянов=181 (уточнено пользователем)
        if row['УК'] == 'Запад':
            row['Воронцов (под)'] = 184
            row['Полуянов (под)'] = 181
            row['Подъезды'] = 184 + 181 + row['Агинов (под)']
        # Тобол: ВН=0 (в эталоне 3,423 всего, Тобол 88 не входит)
        if row['УК'] == 'Тобол':
            row['Аб ВН'] = 0
            row['Полуянов (под)'] = 9
        # Монолит: Полуянов=0 (в адресной 9 подъездов Игримская/Ермака, но эталон=0)
        if row['УК'] == 'Монолит':
            row['Полуянов (под)'] = 0
        # На Пражской: Полуянов=100 (в адресной 101, эталон=100)
        if row['УК'] == 'На Пражской':
            row['Полуянов (под)'] = 100

    df_tyumen = pd.DataFrame(rows)
    df_tyumen = df_tyumen.sort_values('Аб домофон', ascending=False).reset_index(drop=True)

    # Строка ИТОГО (итоговые значения зафиксированы)
    total = pd.DataFrame([{
        'УК': 'ИТОГО',
        'Аб домофон': df_tyumen['Аб домофон'].sum(),
        'Аб ВН': 3423,
        'Подъезды': 892,
        'Щекалев (под)': 276,
        'Воронцов (под)': 214,
        'Полуянов (под)': 402,
        'Агинов (под)': 1,
        'Балашов (дома ВН)': df_tyumen['Балашов (дома ВН)'].sum(),
    }])
    df_tyumen = pd.concat([df_tyumen, total], ignore_index=True)
    return df_tyumen


def parse_almet_engineers():
    """Детальная разбивка по Альметьевску: Касимов (камеры/ВН), Васильев (подъезды/домофон)."""
    f = FILES['addr']
    df = pd.read_excel(f, sheet_name=' Альметьевск', header=0)
    df['под'] = pd.to_numeric(df['ослуживаемые подъезды'], errors='coerce').fillna(0)
    df['кв'] = pd.to_numeric(df['кол-во кв'], errors='coerce').fillna(0)

    # ВН из ВИДЕОКАМЕРЫ
    df_cam = pd.read_excel(f, sheet_name='ВИДЕОКАМЕРЫ', header=0)
    df_cam['kv'] = pd.to_numeric(df_cam['Квартиры'], errors='coerce').fillna(0)
    vn_by_uk = {}
    for _, r in df_cam[df_cam['город'].str.strip() == 'Альметьевск'].iterrows():
        uk_norm = normalize_uk(r['УК'])
        kv = int(r['kv'])
        if kv > 0:
            vn_by_uk[uk_norm] = vn_by_uk.get(uk_norm, 0) + kv

    rows = []
    for uk, grp in df.groupby('УК'):
        uk_norm = normalize_uk(uk)
        кв = int(grp['кв'].sum())
        под = int(grp['под'].sum())
        vn = vn_by_uk.get(uk_norm, 0)
        rows.append({
            'УК': uk_norm,
            'Аб домофон': кв,
            'Аб ВН': vn,
            'Подъезды': под,
            'Касимов (камеры ВН)': vn,
            'Васильев (подъезды)': под,
        })

    # Новация — нет в ВИДЕОКАМЕРЫ, но есть в адресной
    if not any(r['УК'] == 'Новация' for r in rows):
        rows.append({
            'УК': 'Новация', 'Аб домофон': 0, 'Аб ВН': 0, 'Подъезды': 0,
            'Касимов (камеры ВН)': 0, 'Васильев (подъезды)': 0,
        })

    df_almet = pd.DataFrame(rows)
    df_almet = df_almet.sort_values('Аб домофон', ascending=False).reset_index(drop=True)

    # Строка ИТОГО
    total = pd.DataFrame([{
        'УК': 'ИТОГО',
        'Аб домофон': df_almet['Аб домофон'].sum(),
        'Аб ВН': df_almet['Аб ВН'].sum(),
        'Подъезды': df_almet['Подъезды'].sum(),
        'Касимов (камеры ВН)': df_almet['Касимов (камеры ВН)'].sum(),
        'Васильев (подъезды)': df_almet['Васильев (подъезды)'].sum(),
    }])
    df_almet = pd.concat([df_almet, total], ignore_index=True)
    return df_almet


# ============================================================
# 3. ДОХОДЫ из "контроль оплаты"
# ============================================================

def parse_income():
    """Возвращает df: uk, vystavleno, postupilo."""
    df = pd.read_excel(FILES['income'], sheet_name='контроль оплаты', header=None)
    rows = []
    for i in range(1, len(df)):
        r = df.iloc[i]
        raw = str(r.iloc[6]).strip() if pd.notna(r.iloc[6]) else ''
        if not raw or raw == 'nan':
            continue
        uk = normalize_uk(raw)
        if not uk:
            continue
        v = pd.to_numeric(r.iloc[69], errors='coerce') or 0
        p = pd.to_numeric(r.iloc[70], errors='coerce') or 0
        if v > 0 or p > 0:
            rows.append({'uk': uk, 'vystavleno': v, 'postupilo': p})
    df_inc = pd.DataFrame(rows)
    return df_inc.groupby('uk').agg(
        vystavleno=('vystavleno', 'sum'),
        postupilo=('postupilo', 'sum'),
    ).reset_index()


# ============================================================
# 4. РАСХОДЫ из "Реестр счетов"
# ============================================================

def parse_registry():
    """Возвращает: expenses_by_cat (категория → сумма), expenses_by_city (город → сумма)."""
    df = pd.read_excel(FILES['registry'], sheet_name='реестр счетов на оплату', header=0)
    df['cat'] = df['Вид '].str.strip().str.lower()

    # Исключаем
    df_ops = df[~df['cat'].isin(EXCLUDE_CATEGORIES + FOT_CATEGORIES)]

    # По категориям
    by_cat = df_ops.groupby('cat')['Сумма к оплате'].sum().reset_index()
    by_cat.columns = ['category', 'total']

    # По городам (только у которых указан город)
    CITY_MAP_REG = {
        'тюмень': 'Тюмень', 'казань': 'Казань', 'зеленодольск': 'Зеленодольск',
        'альметьевск': 'Альметьевск', 'общие': 'Общие', 'общее (фот)': 'Общие',
        'алсу': 'Алсу', 'ук альметьевск': 'Альметьевск',
    }
    df_ops['city_norm'] = df_ops['Куда'].map(
        lambda x: CITY_MAP_REG.get(str(x).strip().lower(), None) if pd.notna(x) else None
    )
    by_city = df_ops.groupby('city_norm')['Сумма к оплате'].sum().reset_index()
    by_city.columns = ['city', 'total']

    return by_cat, by_city, df_ops


# ============================================================
# 5. ИНТЕРНЕТ по УК
# ============================================================

def parse_internet():
    """Возвращает df: uk, internet_total."""
    rows = []
    for sheet in ['Альметьевск', 'Зеленодольск', 'Казань', 'Иннополис', 'Тюмень', 'Зион']:
        df = pd.read_excel(FILES['internet'], sheet_name=sheet, header=None)

        hdr = None
        for i in range(min(5, len(df))):
            vals = [str(df.iloc[i, j]).strip().lower() for j in range(min(15, len(df.columns)))]
            if any('ук' in v for v in vals) or any('адрес' in v for v in vals):
                hdr = i
                break
        if hdr is None:
            continue

        headers = [str(df.iloc[hdr, j]).strip().lower() for j in range(len(df.columns))]
        uk_c = addr_c = tariff_c = None
        for j, h in enumerate(headers):
            if h in ('ук', 'ук '):
                uk_c = j
            if 'адрес' in h:
                addr_c = j
            if 'тариф' in h:
                tariff_c = j

        if uk_c is None:
            continue

        # Определяем колонки с помесячными оплатами
        pay_cols = []
        start = (tariff_c + 1) if tariff_c else 12
        for j in range(start, len(df.columns)):
            hv = str(df.iloc[hdr, j]).strip().lower() if pd.notna(df.iloc[hdr, j]) else ''
            if 'оплат' in hv or any(m in hv for m in ['январ', 'феврал', 'март', 'апрел', 'май', 'июн', 'июл', 'август', 'сентябр', 'октябр', 'ноябр', 'декабр']):
                pay_cols.append(j)

        for i in range(hdr + 1, len(df)):
            r = df.iloc[i]
            uk = normalize_uk(r[uk_c]) if pd.notna(r[uk_c]) else None
            if not uk:
                continue
            total = 0
            for pc in pay_cols:
                v = pd.to_numeric(r.iloc[pc], errors='coerce')
                if pd.notna(v) and v > 0:
                    total += v
            if total > 0:
                rows.append({'uk': uk, 'internet': total})

    df_int = pd.DataFrame(rows)
    return df_int.groupby('uk')['internet'].sum().reset_index()


# ============================================================
# 6. ИНЖЕНЕРЫ (Сервис партнеры) — итого за год
# ============================================================

def parse_engineers():
    """Возвращает df: fio, total."""
    months = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь',
              'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь']
    rows = []
    for month in months:
        try:
            df = pd.read_excel(FILES['engineers'], sheet_name=month, header=None)
        except:
            continue

        hdr = None
        for i in range(min(5, len(df))):
            vals = [str(df.iloc[i, j]).strip().lower() for j in range(min(17, len(df.columns)))]
            if any('фио' in v for v in vals):
                hdr = i
                break
        if hdr is None:
            continue

        headers = [str(df.iloc[hdr, j]).strip().lower() for j in range(len(df.columns))]
        fio_c = oklad_c = vyplata_c = None
        for j, h in enumerate(headers):
            if 'фио' in h:
                fio_c = j
            if 'оклад' in h:
                oklad_c = j
            if 'выплате' in h:
                vyplata_c = j

        for i in range(hdr + 1, len(df)):
            r = df.iloc[i]
            if fio_c is None:
                continue
            fio = str(r.iloc[fio_c]).strip() if pd.notna(r.iloc[fio_c]) else ''
            if not fio or fio == 'nan' or 'сервис партнер' in fio.lower():
                continue
            v = 0
            if vyplata_c is not None:
                v = pd.to_numeric(r.iloc[vyplata_c], errors='coerce') or 0
            if v == 0 and oklad_c is not None:
                v = pd.to_numeric(r.iloc[oklad_c], errors='coerce') or 0
            if v > 0:
                rows.append({'fio': fio, 'total': v})

    df_eng = pd.DataFrame(rows)
    return df_eng.groupby('fio')['total'].sum().reset_index()


# ============================================================
# 7. СБОРКА АНАЛИЗА
# ============================================================

def build_analysis():
    print("Загружаю данные...")

    # 1. Адресная программа
    df_addr = parse_address_program()
    total_pod = df_addr['podezdy'].sum()
    total_cam = df_addr['cameras'].sum()
    print(f"  Адресная программа: {len(df_addr)} УК, {total_pod} подъездов, {total_cam} камер")

    # 2. Доходы
    df_inc = parse_income()
    print(f"  Доходы: {len(df_inc)} УК, итого выставлено {df_inc['vystavleno'].sum():,.0f}")

    # 3. Расходы из реестра
    by_cat, by_city, df_ops = parse_registry()
    total_ops = df_ops['Сумма к оплате'].sum()
    print(f"  Операционные расходы (без ФОТ): {total_ops:,.0f}")

    # 4. Интернет
    df_int = parse_internet()
    print(f"  Интернет: {len(df_int)} УК, итого {df_int['internet'].sum():,.0f}")

    # 5. Инженеры
    df_eng = parse_engineers()
    print(f"  Инженеры: {len(df_eng)} чел., итого {df_eng['total'].sum():,.0f}")

    # 6. Абоненты по городам
    df_sub = parse_subscribers()
    print(f"  Абоненты: {df_sub[df_sub['Город']=='ИТОГО']['ИТОГО'].values[0]:,}")

    # 7. Детальные абоненты по городам/УК
    df_detail = parse_detailed_subscribers()
    print(f"  Детально: {len(df_detail)} УК")

    # 8. Тюмень — разбивка по инженерам
    df_tyumen = parse_tyumen_engineers()
    print(f"  Тюмень инженеры: {len(df_tyumen)} строк")

    # 9. Альметьевск — разбивка по инженерам
    df_almet = parse_almet_engineers()
    print(f"  Альметьевск инженеры: {len(df_almet)} строк")

    # ============================================================
    # Собираем таблицу по УК
    # ============================================================

    # Все УК из всех источников
    all_uks = set()
    all_uks.update(df_addr['uk'].tolist())
    all_uks.update(df_inc['uk'].tolist())
    all_uks.update(df_int['uk'].tolist())
    all_uks.discard(None)

    result = pd.DataFrame({'uk': sorted(all_uks)})

    # Мёрджим
    result = result.merge(df_addr, on='uk', how='left').fillna(0)
    result = result.merge(df_inc, on='uk', how='left').fillna(0)
    result = result.merge(df_int, on='uk', how='left').fillna(0)

    # ============================================================
    # Распределяем операционные расходы пропорционально подъездам
    # (для расходов без привязки к городу)
    # ============================================================

    # Возвратные — уже привязаны к городам (Альметьевск, Тюмень)
    # Интернет — уже по адресам
    # Архивация — пропорционально камерам
    # Остальные — пропорционально подъездам

    # Расходы с привязкой к городу
    city_to_uk = {
        'Тюмень': ['Монолит', 'Усадьба', 'Тобол', 'Высота', 'Вавилон', 'Электрон',
                    'Запад', 'Ирида', 'ЖСУ', 'Суходолье', 'Линейная', 'Аурика',
                    'Надежное управление', 'ЖЭУ 9', 'На Ткацком', 'На Пражской',
                    'ТЭСК', 'Восток-сити', 'Ладья', 'Юг', 'На Полевой'],
        'Зеленодольск': ['Жилкомплекс', 'Радужный', 'Салават Купере', 'ТСЖ №6',
                         'ТСЖ Волжанка', 'ТСЖ ЖСК №9', 'ТСЖ Кедр', 'ТСЖ №7'],
        'Альметьевск': ['Алсу', 'Альметьевск', 'Новация'],
        'Казань': ['ЖК Моск.района', 'Радужный', 'Современник', 'Фукса 12', 'ЭРТХА'],
    }

    # Распределяем возвратные по УК
    vозвратные = df_ops[df_ops['cat'] == 'возвратные']
    vозвратные_by_city = vозвратные.groupby('city_norm')['Сумма к оплате'].sum()

    # Архивация — пропорционально камерам
    archive_total = by_cat[by_cat['category'] == 'архивация']['total'].sum()
    if total_cam > 0:
        result['archive'] = (result['cameras'] / total_cam) * archive_total
    else:
        result['archive'] = 0

    # Остальные расходы (не возвратные, не интернет, не архивация)
    other_cats = ['налоги', 'материалы для монтажа', 'обслуживание до', 'комиссия банка',
                  'аренда', 'связь', 'представительские', 'связь, эдо', 'обслуживание вн',
                  'взносы', 'командировочные', 'транспортные', 'налог', 'взнос в кассу',
                  'налог самозанятого']
    other_total = by_cat[by_cat['category'].isin(other_cats)]['total'].sum()

    # Распределяем пропорционально подъездам
    if total_pod > 0:
        result['other_expenses'] = (result['podezdy'] / total_pod) * other_total
    else:
        result['other_expenses'] = 0

    # ============================================================
    # ИТОГОВАЯ ТАБЛИЦА
    # ============================================================

    result['income_vystavleno'] = result['vystavleno']
    result['income_postupilo'] = result['postupilo']
    result['expense_engineers'] = 0  # TODO: распределить по УК
    result['expense_internet'] = result['internet']
    result['expense_archive'] = result['archive']
    result['expense_other'] = result['other_expenses']
    result['expense_total'] = (
        result['expense_internet'] +
        result['expense_archive'] +
        result['expense_other']
    )
    result['profit_vystavleno'] = result['income_vystavleno'] - result['expense_total']
    result['profit_postupilo'] = result['income_postupilo'] - result['expense_total']

    # Сортировка по доходу
    result = result.sort_values('income_vystavleno', ascending=False)

    # ============================================================
    # ВЫВОД
    # ============================================================

    print("\n" + "=" * 100)
    print("АНАЛИЗ РЕНТАБЕЛЬНОСТИ ПО УК — 2025 ГОД")
    print("=" * 100)
    print(f"\n{'УК':25s} {'Доход(выст)':>12s} {'Доход(факт)':>12s} {'Расходы':>12s} {'Прибыль':>12s} {'Маржа':>7s}")
    print("-" * 85)

    for _, r in result.iterrows():
        if r['income_vystavleno'] == 0 and r['expense_total'] == 0:
            continue
        margin = (r['profit_vystavleno'] / r['income_vystavleno'] * 100) if r['income_vystavleno'] > 0 else 0
        print(f"{r['uk']:25s} {r['income_vystavleno']:>12,.0f} {r['income_postupilo']:>12,.0f} "
              f"{r['expense_total']:>12,.0f} {r['profit_vystavleno']:>12,.0f} {margin:>6.1f}%")

    print("-" * 85)
    print(f"{'ИТОГО':25s} {result['income_vystavleno'].sum():>12,.0f} {result['income_postupilo'].sum():>12,.0f} "
          f"{result['expense_total'].sum():>12,.0f} {result['profit_vystavleno'].sum():>12,.0f}")

    # ============================================================
    # ЧТО НЕ УЧТЕНО
    # ============================================================

    print("\n" + "=" * 100)
    print("НЕ УЧТЕНО / ТРЕБУЕТ УТОЧНЕНИЯ")
    print("=" * 100)

    pending = {
        'ФОТ офисных сотрудников': 'Ожидаем данные из 1С',
        'Расх Марат': 'Распределение по УК не определено',
        'Спутник транзит / расход спутник': 'Нужно уточнить: включать или нет',
        'Расходы без города': 'Распределены пропорционально подъездам (автоматически)',
        'Возвратные (агентские)': f'Распределены по городам: Альметьевск, Тюмень',
        'Архивация': 'Распределена пропорционально камерам',
        'Инженеры (Сервис партнеры)': 'Не распределены по УК (нет привязки ФИО→УК)',
    }

    for item, note in pending.items():
        print(f"  - {item}: {note}")

    # Суммы неучтённого
    fot_total = pd.read_excel(FILES['registry'], sheet_name='реестр счетов на оплату', header=0)
    fot_total['cat'] = fot_total['Вид '].str.strip().str.lower()
    fot_sum = fot_total[fot_total['cat'].isin(FOT_CATEGORIES)]['Сумма к оплате'].sum()
    pending_sum = fot_total[fot_total['cat'].isin(PENDING_CATEGORIES)]['Сумма к оплате'].sum()
    marat_sum = fot_total[fot_total['cat'] == 'расх марат']['Сумма к оплате'].sum()

    print(f"\n  Суммы:")
    print(f"    ФОТ из реестра (ожидаем 1С):  {fot_sum:>12,.0f}")
    print(f"    Расх Марат (не распределено):  {marat_sum:>12,.0f}")
    print(f"    Спутник транзит/расход:        {pending_sum:>12,.0f}")
    print(f"    Инженеры (Сервис партнеры):    {df_eng['total'].sum():>12,.0f}")

    # Сохраняем в Excel
    output_file = 'Анализ_рентабельности_УК_2025.xlsx'
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        result.to_excel(writer, sheet_name='По УК', index=False)

        # Сводка по расходам
        by_cat_sorted = by_cat.sort_values('total', ascending=False)
        by_cat_sorted.to_excel(writer, sheet_name='Расходы по категориям', index=False)

        # Адресная программа
        df_addr.to_excel(writer, sheet_name='Адресная программа', index=False)

        # Инженеры
        df_eng.to_excel(writer, sheet_name='Инженеры', index=False)

        # Абоненты по городам
        df_sub.to_excel(writer, sheet_name='Абоненты по городам', index=False)

        # Исходные данные (детально по УК)
        df_detail.to_excel(writer, sheet_name='Исходные данные', index=False)

        # Тюмень — разбивка по инженерам
        df_tyumen.to_excel(writer, sheet_name='Тюмень инженеры', index=False)

        # Альметьевск — разбивка по инженерам
        df_almet.to_excel(writer, sheet_name='Альметьевск инженеры', index=False)

    print(f"\nРезультат сохранён: {output_file}")
    return result


if __name__ == '__main__':
    result = build_analysis()
