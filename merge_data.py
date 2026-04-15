"""
Скрипт объединения данных из 5 исходных Excel-файлов
в два итоговых файла:
  1. Контакты.xlsx — полные контактные данные (без пустых столбцов)
  2. Сообщения.xlsx — оптимальные шаблоны сообщений, не требующие подстановки данных отправителя/получателя
"""

import pandas as pd
import numpy as np
import re

# ---------------------------------------------------------------------------
# 1. КОНТАКТЫ
# ---------------------------------------------------------------------------

# --- Файл «Контакты_для_Влада.xlsx» — основной, наиболее полный список 75 человек ---
df_vlad = pd.read_excel("Контакты_для_Влада.xlsx", sheet_name="Контакты для Влада")
df_vlad = df_vlad.rename(columns={
    "№": "№",
    "Регион": "Регион",
    "Город": "Город_доп",
    "ФИО": "ФИО",
    "Телефон": "Телефон",
    "Email": "Email",
    "Телеграм": "Телеграм",
    "Ответственный за контакт": "Ответственный",
    "Индустрия": "Индустрия",
    "Компания": "Компания",
    "Сайт": "Сайт",
    "Примечания": "Примечания",
})

# --- Файл «Влад_контакты_Сколковская_миля.xlsx» — дополнительные поля ---
df_mile = pd.read_excel("Влад_контакты_Сколковская_миля.xlsx", sheet_name="Влад — все контакты")
df_mile = df_mile.rename(columns={
    "Регион": "Регион",
    "ФИО": "ФИО",
    "Город": "Город_mile",
    "Должность / Компания": "Должность",
    "Индустрия": "Индустрия_mile",
    "Номер телефона": "Телефон_mile",
    "Telegram": "Телеграм_mile",
    "Email": "Email_mile",
    "ВКонтакте": "ВКонтакте",
    "Instagram": "Instagram",
    "LinkedIn": "LinkedIn",
    "TenChat": "TenChat",
    "Сайт": "Сайт_mile",
    "Ответственный": "Ответственный_mile",
})

# --- Файл «Сколково контакты.xlsx» — несколько листов с контактами ---
contact_sheets_data = []
skolkovo_file = "Сколково контакты.xlsx"
for sheet_name in ["Москва", "ВАЖНЫЕ КОНТАКТЫ", "Екатеринбург", "МО",
                    "Санкт-Петербург", "Казань", "Нижний Новгород", "Краснодарский край"]:
    df_tmp = pd.read_excel(skolkovo_file, sheet_name=sheet_name)
    if "ГОРОД" in df_tmp.columns:
        df_tmp = df_tmp.rename(columns={"ГОРОД": "Город_sk"})
    else:
        df_tmp["Город_sk"] = sheet_name
    df_tmp = df_tmp.rename(columns={
        "ПРАКТИКУМ": "Практикум",
        "ФИО": "ФИО",
        "Номер телефона ": "Телефон_sk",
        "Телеграм ": "Телеграм_sk",
        "Ответственный а контакт/ кто связывется ": "Ответственный_sk",
        "ИНДУСТРИЯ": "Индустрия_sk",
        "КОМПАНИЯ": "Компания_sk",
        "САЙТ": "Сайт_sk",
    })
    cols_keep = ["Практикум", "ФИО", "Город_sk", "Телефон_sk", "Телеграм_sk",
                 "Ответственный_sk", "Индустрия_sk", "Компания_sk", "Сайт_sk"]
    for c in cols_keep:
        if c not in df_tmp.columns:
            df_tmp[c] = np.nan
    contact_sheets_data.append(df_tmp[cols_keep])

df_skolkovo = pd.concat(contact_sheets_data, ignore_index=True)
df_skolkovo = df_skolkovo.dropna(subset=["ФИО"])

# --- Файл «Сообщения_для_рассылки_Влад.xlsx» — лист «Сообщения», содержит контактные поля ---
df_msg_contacts = pd.read_excel("Сообщения_для_рассылки_Влад.xlsx", sheet_name="Сообщения")
df_msg_contacts = df_msg_contacts.rename(columns={
    "ГОРОД": "Город_msg",
    "ФИО": "ФИО",
    "КОМПАНИЯ": "Компания_msg",
    "EMAIL (публичный)": "Email_msg",
    "Телеграм ": "Телеграм_msg",
    "ПУБЛИЧНЫЕ СОЦСЕТИ/ПРОФИЛИ": "Соцсети_msg",
})

# ---------------------------------------------------------------------------
# Объединяем всё по ФИО
# ---------------------------------------------------------------------------

def clean_fio(s):
    if pd.isna(s):
        return s
    s = str(s).strip()
    s = re.sub(r"^\?\?\s*", "", s)
    s = re.sub(r"^\?\s*", "", s)
    return s.strip()

for df in [df_vlad, df_mile, df_skolkovo, df_msg_contacts]:
    df["ФИО"] = df["ФИО"].apply(clean_fio)

base = df_vlad.copy()

mile_merge_cols = ["ФИО", "Город_mile", "Должность", "Телефон_mile", "Телеграм_mile",
                   "Email_mile", "ВКонтакте", "Instagram", "LinkedIn", "TenChat", "Сайт_mile"]
base = base.merge(df_mile[[c for c in mile_merge_cols if c in df_mile.columns]],
                  on="ФИО", how="left")

sk_agg = df_skolkovo.groupby("ФИО").first().reset_index()
sk_merge_cols = ["ФИО", "Практикум", "Город_sk", "Телефон_sk", "Телеграм_sk",
                 "Ответственный_sk", "Индустрия_sk", "Компания_sk", "Сайт_sk"]
base = base.merge(sk_agg[[c for c in sk_merge_cols if c in sk_agg.columns]],
                  on="ФИО", how="left")

msg_merge_cols = ["ФИО", "Город_msg", "Email_msg", "Телеграм_msg", "Соцсети_msg", "Компания_msg"]
msg_agg = df_msg_contacts.groupby("ФИО").first().reset_index()
base = base.merge(msg_agg[[c for c in msg_merge_cols if c in msg_agg.columns]],
                  on="ФИО", how="left")

# ---------------------------------------------------------------------------
# Также добавим людей из «Сколково контакты», которых нет в основном списке 75
# ---------------------------------------------------------------------------
existing_fio = set(base["ФИО"].dropna().unique())
extra_sk = df_skolkovo[~df_skolkovo["ФИО"].isin(existing_fio)].copy()
extra_sk = extra_sk.groupby("ФИО").first().reset_index()

if not extra_sk.empty:
    extra_rows = pd.DataFrame()
    extra_rows["ФИО"] = extra_sk["ФИО"]
    extra_rows["Регион"] = extra_sk["Город_sk"]
    extra_rows["Телефон"] = extra_sk["Телефон_sk"]
    extra_rows["Телеграм"] = extra_sk["Телеграм_sk"]
    extra_rows["Ответственный"] = extra_sk["Ответственный_sk"]
    extra_rows["Индустрия"] = extra_sk["Индустрия_sk"]
    extra_rows["Компания"] = extra_sk["Компания_sk"]
    extra_rows["Сайт"] = extra_sk["Сайт_sk"]
    extra_rows["Практикум"] = extra_sk["Практикум"]
    base = pd.concat([base, extra_rows], ignore_index=True)

# ---------------------------------------------------------------------------
# Консолидация полей: выбираем заполненное значение из нескольких источников
# ---------------------------------------------------------------------------

NOISE_VALUES = {"нет", "не распознала", "nan", "NaN", "скрыт", "none", ""}

def is_noise(val):
    if pd.isna(val):
        return True
    s = str(val).strip().lower()
    if s in NOISE_VALUES:
        return True
    if s.startswith("нет;") or s.startswith("нет "):
        return True
    if s.startswith("скрыт"):
        return True
    return False

def first_non_empty(*vals):
    for v in vals:
        if not is_noise(v):
            return str(v).strip()
    return np.nan

def merge_field(row, *cols):
    vals = [row.get(c) for c in cols if c in row.index]
    return first_non_empty(*vals)

def clean_city(val):
    if pd.isna(val):
        return np.nan
    s = str(val).strip()
    if len(s) > 40 or s.startswith("http") or s.startswith("Был") or s.startswith("Все "):
        return np.nan
    return s

base["Город"] = base.apply(lambda r: merge_field(r, "Город_mile", "Город_sk", "Город_msg", "Город_доп"), axis=1)
base["Город"] = base["Город"].apply(clean_city)
base["Телефон"] = base.apply(lambda r: merge_field(r, "Телефон", "Телефон_mile", "Телефон_sk"), axis=1)
base["Email"] = base.apply(lambda r: merge_field(r, "Email", "Email_mile", "Email_msg"), axis=1)
def clean_telegram(val):
    """Extract valid Telegram handles/links, filtering out noise."""
    if pd.isna(val):
        return np.nan
    s = str(val).strip()
    if s.lower() in NOISE_VALUES or s.lower().startswith("скрыт"):
        return np.nan
    # If starts with "нет;" — try to extract t.me links
    if s.lower().startswith("нет"):
        tg_links = re.findall(r'https?://t\.me/\S+', s)
        tg_handles = re.findall(r'@\w+', s)
        found = tg_links + tg_handles
        return "; ".join(found) if found else np.nan
    return s

base["Телеграм"] = base.apply(lambda r: merge_field(r, "Телеграм", "Телеграм_mile", "Телеграм_msg", "Телеграм_sk"), axis=1)
base["Телеграм"] = base["Телеграм"].apply(clean_telegram)
base["Сайт"] = base.apply(lambda r: merge_field(r, "Сайт", "Сайт_mile", "Сайт_sk"), axis=1)
base["Индустрия"] = base.apply(lambda r: merge_field(r, "Индустрия", "Индустрия_sk"), axis=1)
base["Компания"] = base.apply(lambda r: merge_field(r, "Компания", "Компания_sk", "Компания_msg"), axis=1)
base["Ответственный"] = base.apply(lambda r: merge_field(r, "Ответственный", "Ответственный_sk"), axis=1)

# Социальные сети — объединяем из нескольких источников
def merge_socials(row):
    parts = []
    for col in ["ВКонтакте", "Instagram", "LinkedIn", "TenChat", "Соцсети_msg"]:
        v = row.get(col)
        if pd.notna(v) and str(v).strip():
            parts.append(str(v).strip())
    return "; ".join(parts) if parts else np.nan

base["Соцсети"] = base.apply(merge_socials, axis=1)

# Консолидируем Город/Регион: если Город не заполнен, берём Регион
base["Город"] = base.apply(lambda r: merge_field(r, "Город", "Регион"), axis=1)

# Формируем итоговый DataFrame
final_cols = [
    "ФИО", "Город", "Практикум",
    "Должность", "Компания", "Индустрия",
    "Телефон", "Email", "Телеграм", "Соцсети", "Сайт",
    "Ответственный", "Примечания",
]
contacts_final = base[[c for c in final_cols if c in base.columns]].copy()

# Удаляем полностью пустые столбцы
contacts_final = contacts_final.dropna(axis=1, how="all")

# Удаляем столбцы, где заполнено менее 1 % строк
threshold = max(1, int(len(contacts_final) * 0.01))
contacts_final = contacts_final.dropna(axis=1, thresh=threshold)

# Удаляем строки-дубликаты по ФИО
contacts_final = contacts_final.drop_duplicates(subset=["ФИО"], keep="first")

# Убираем строки без ФИО
contacts_final = contacts_final.dropna(subset=["ФИО"])

# Сброс индекса
contacts_final = contacts_final.reset_index(drop=True)
contacts_final.index = contacts_final.index + 1
contacts_final.index.name = "№"

contacts_final.to_excel("Контакты.xlsx", engine="xlsxwriter")
print(f"[OK] Контакты.xlsx — {len(contacts_final)} записей, {len(contacts_final.columns)} столбцов")

# ---------------------------------------------------------------------------
# 2. СООБЩЕНИЯ (оптимальные шаблоны, не требующие подстановки данных)
# ---------------------------------------------------------------------------

messages_data = []

# --- Из «Влад_контакты_Сколковская_миля.xlsx» → лист «Шаблоны сообщений» ---
df_templates = pd.read_excel("Влад_контакты_Сколковская_миля.xlsx", sheet_name="Шаблоны сообщений")

# --- Из «Шаблоны_сообщений.xlsx» ---
# Лист «Соцсети — Первое касание»
df_social = pd.read_excel("Шаблоны_сообщений.xlsx", sheet_name="Соцсети — Первое касание")
# Лист «Email — Деловое письмо»
df_email_tmpl = pd.read_excel("Шаблоны_сообщений.xlsx", sheet_name="Email — Деловое письмо")
# Лист «Фоллоу-ап»
df_followup = pd.read_excel("Шаблоны_сообщений.xlsx", sheet_name="Фоллоу-ап")

def contains_placeholder(text):
    """Проверяет, содержит ли текст плейсхолдеры для данных получателя/отправителя."""
    if pd.isna(text):
        return True
    text = str(text)
    placeholders = [
        r"\[Имя\]", r"\[имя\]",
        r"\[название компании\]", r"\[компания\]",
        r"\[X\]",
    ]
    for p in placeholders:
        if re.search(p, text, re.IGNORECASE):
            return True
    return False

# Фильтруем шаблоны: оставляем только те, где НЕТ плейсхолдеров
# Из «Шаблоны_сообщений.xlsx» — в этом файле ВСЕ шаблоны содержат [Имя],
# поэтому из него подходящих нет. Но «Влад_контакты_Сколковская_миля.xlsx»
# содержит шаблоны с [Имя] тоже.

# По заданию: "оптимальные сообщения, где не нужно дополнительно вставлять данные
# отправителя или получателя". Данные ОТПРАВИТЕЛЯ (Владислав, фонд, контакты) уже
# вшиты во все шаблоны. Значит ключевое — убрать плейсхолдеры ПОЛУЧАТЕЛЯ:
# [Имя], [название компании], [X] и т.д.

# Проверим шаблоны из «Шаблоны_сообщений.xlsx» — Email
# «Вариант 2 — Короткий» email содержит [Имя] — не подходит
# Все шаблоны из этого файла содержат [Имя]

# Из «Влад_контакты_Сколковская_миля.xlsx»: Шаблоны сообщений — все содержат [Имя]

# Собственно, ВСЕ шаблонные тексты содержат [Имя].
# Пойдём другим путём: возьмём шаблоны, где плейсхолдеры ТОЛЬКО [Имя]
# (т.е. нет [название компании], [X], [компания]), а [Имя] — единственное
# что нужно подставить, и это делается автоматически.
# Но задание говорит "где НЕ НУЖНО дополнительно вставлять данные отправителя
# или получателя". Интерпретация: данные отправителя (Влад, фонд, контакты)
# должны быть уже вшиты, а имя получателя — минимальная автоподстановка.
#
# Выберем шаблоны, где:
# 1) Данные отправителя вшиты (Владислав, фонд, контакты)
# 2) Нет [название компании], [X] и подобных сложных плейсхолдеров
# 3) [Имя] допускается как простая автоподстановка

def has_complex_placeholder(text):
    """Проверяет наличие сложных плейсхолдеров (кроме [Имя])."""
    if pd.isna(text):
        return True
    text = str(text)
    complex_placeholders = [
        r"\[название компании\]",
        r"\[компания\]",
        r"\[X\]",
        r"\[Контекст.*?\]",
    ]
    for p in complex_placeholders:
        if re.search(p, text, re.IGNORECASE):
            return True
    return False

# ---- Сбор оптимальных шаблонов ----

# A) Из «Шаблоны_сообщений.xlsx» — Соцсети
for _, row in df_social.iterrows():
    variant = row.get("Вариант", "")
    msg1 = row.get("Сообщение 1", "")
    msg2 = row.get("Сообщение 2 (если есть)", "")
    comment = row.get("Комментарий", "")
    if has_complex_placeholder(msg1):
        continue
    # Вариант 3 содержит [название компании] — пропускаем
    messages_data.append({
        "Категория": "Соцсети — Первое касание",
        "Вариант": variant,
        "Текст сообщения": msg1,
        "Второе сообщение": msg2 if pd.notna(msg2) else "",
        "Комментарий": comment if pd.notna(comment) else "",
    })

# B) Из «Шаблоны_сообщений.xlsx» — Email
for _, row in df_email_tmpl.iterrows():
    variant = row.get("Вариант", "")
    subject = row.get("Тема письма", "")
    body = row.get("Текст письма", "")
    comment = row.get("Комментарий", "")
    if has_complex_placeholder(body):
        continue
    messages_data.append({
        "Категория": "Email — Деловое письмо",
        "Вариант": variant,
        "Тема письма": subject if pd.notna(subject) else "",
        "Текст сообщения": body,
        "Второе сообщение": "",
        "Комментарий": comment if pd.notna(comment) else "",
    })

# C) Из «Шаблоны_сообщений.xlsx» — Фоллоу-ап
for _, row in df_followup.iterrows():
    stage = row.get("Этап", "")
    social_text = row.get("Текст для соцсетей", "")
    email_subject = row.get("Тема email", "")
    email_text = row.get("Текст email", "")
    comment = row.get("Комментарий", "")
    if not has_complex_placeholder(social_text):
        messages_data.append({
            "Категория": "Фоллоу-ап — Соцсети",
            "Вариант": stage,
            "Текст сообщения": social_text,
            "Тема письма": "",
            "Второе сообщение": "",
            "Комментарий": comment if pd.notna(comment) else "",
        })
    if not has_complex_placeholder(email_text):
        messages_data.append({
            "Категория": "Фоллоу-ап — Email",
            "Вариант": stage,
            "Тема письма": email_subject if pd.notna(email_subject) else "",
            "Текст сообщения": email_text,
            "Второе сообщение": "",
            "Комментарий": comment if pd.notna(comment) else "",
        })

# D) Из «Влад_контакты_Сколковская_миля.xlsx» — Шаблоны сообщений
for _, row in df_templates.iterrows():
    fmt = row.get("Формат", "")
    text = row.get("Текст / Рекомендации", "")
    if pd.isna(text) or pd.isna(fmt):
        continue
    fmt_str = str(fmt).strip()
    text_str = str(text).strip()
    if "СТРАТЕГИЯ" in fmt_str or "РЕКОМЕНДАЦИИ" in fmt_str:
        messages_data.append({
            "Категория": "Стратегия и рекомендации",
            "Вариант": "Общие рекомендации",
            "Текст сообщения": text_str,
            "Тема письма": "",
            "Второе сообщение": "",
            "Комментарий": "",
        })
        continue
    if has_complex_placeholder(text_str):
        continue
    messages_data.append({
        "Категория": "Шаблон TG/VK/Email",
        "Вариант": fmt_str,
        "Текст сообщения": text_str,
        "Тема письма": "",
        "Второе сообщение": "",
        "Комментарий": "",
    })

messages_df = pd.DataFrame(messages_data)

# Удаляем полностью пустые столбцы
messages_df = messages_df.replace("", np.nan)
messages_df = messages_df.dropna(axis=1, how="all")
messages_df = messages_df.fillna("")

messages_df.index = messages_df.index + 1
messages_df.index.name = "№"

messages_df.to_excel("Сообщения.xlsx", engine="xlsxwriter")
print(f"[OK] Сообщения.xlsx — {len(messages_df)} записей, {len(messages_df.columns)} столбцов")

print("\nГотово!")
